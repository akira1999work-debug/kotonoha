"""固有名詞の訂正精度 評価ハーネス.

SAPI (Microsoft Haruka) で読み上げた音声を Whisper -> LLM 整形に通し、
正解表記が最終出力に現れるかを測る。認識層と整形層を分けて集計するので、
「どちらを直すべきか」が毎回わかる。

使い方:
    .venv\\Scripts\\python.exe eval\\run_eval.py                # 通常実行
    .venv\\Scripts\\python.exe eval\\run_eval.py --label 施策名  # 結果に名前を付けて保存
    .venv\\Scripts\\python.exe eval\\run_eval.py --compare       # 保存済み結果を並べる

注意: TTS 音声は実際の発話とは癖が違う。ここのスコアは施策の相対比較に使うもので、
絶対的な実力値ではない。
"""

from __future__ import annotations

import argparse
import io
import json
import contextlib
import sys
import time
from pathlib import Path

EVAL_DIR = Path(__file__).parent
REPO_DIR = EVAL_DIR.parent
AUDIO_DIR = EVAL_DIR / "audio"
RESULTS_DIR = EVAL_DIR / "results"

sys.path.insert(0, str(REPO_DIR))


def ensure_audio(cases: list[dict]) -> None:
    """未生成の case だけ SAPI で wav を作る。既存はそのまま使う(比較の再現性のため)。"""
    AUDIO_DIR.mkdir(parents=True, exist_ok=True)
    missing = [c for c in cases if not (AUDIO_DIR / f"{c['id']}.wav").exists()]
    if not missing:
        return

    import win32com.client

    voice = win32com.client.Dispatch("SAPI.SpVoice")
    for v in voice.GetVoices():
        if "Haruka" in v.GetDescription() or "ja-JP" in v.GetDescription():
            voice.Voice = v
            break
    else:
        print("[eval] 日本語 SAPI 音声が見つからない。英語音声で代用します", flush=True)

    for c in missing:
        path = AUDIO_DIR / f"{c['id']}.wav"
        stream = win32com.client.Dispatch("SAPI.SpFileStream")
        stream.Open(str(path), 3)  # 3 = SSFMCreateForWrite
        voice.AudioOutputStream = stream
        voice.Speak(c["speak"])
        stream.Close()
        print(f"[eval] 音声生成: {path.name} <- 「{c['speak']}」", flush=True)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--label", default="baseline", help="この実行につける名前")
    ap.add_argument("--compare", action="store_true", help="保存済み結果を一覧表示して終了")
    ap.add_argument(
        "--prompts",
        help="整形プロンプトを差し替える JSON (default/code/casual をキーに持つ)。"
        "config.json を触らずに変種を試すため",
    )
    ap.add_argument("--llm-model", help="LLM モデルを一時的に差し替える (例: qwen2.5:14b-instruct-q4_K_M)")
    ap.add_argument("--no-replace", action="store_true", help="決定的置換を無効にして測る")
    args = ap.parse_args()

    RESULTS_DIR.mkdir(parents=True, exist_ok=True)

    if args.compare:
        rows = []
        for f in sorted(RESULTS_DIR.glob("*.json")):
            d = json.loads(f.read_text(encoding="utf-8"))
            rows.append(d)
        if not rows:
            print("保存済み結果なし")
            return 0
        print(f"{'施策':<26} {'認識層':>8} {'最終':>8} {'忠実度':>8} {'切詰':>6}")
        print("-" * 62)
        for d in rows:
            sim = d.get("mean_similarity", "-")
            tr = d.get("truncations", "-")
            print(
                f"{d['label']:<26} {d['asr_score']:>4}/{d['total']:<3} "
                f"{d['final_score']:>4}/{d['total']:<3} {str(sim):>8} {str(tr):>6}"
            )
        return 0

    cases = json.loads((EVAL_DIR / "cases.json").read_text(encoding="utf-8"))["cases"]
    ensure_audio(cases)

    import voice_input as vi
    from faster_whisper.audio import decode_audio

    cfg = vi.load_config()

    if args.prompts:
        # 差し替えは辞書注入の「後」に行う必要がある (load_config が置換マップを
        # 各プロンプト末尾に append しているため)。同じマップを変種にも付け直す。
        terms = vi.load_user_dictionary()
        hint = vi.build_grouped_llm_hint(vi.score_and_filter(terms, vi.LLM_DICT_TOP_N))
        variants = json.loads(Path(args.prompts).read_text(encoding="utf-8"))
        for key, body in variants.items():
            cfg["prompts"][key] = body + hint
        print(f"[eval] プロンプト差し替え: {args.prompts} ({', '.join(variants)})", flush=True)

    if args.llm_model:
        cfg["llm"]["model"] = args.llm_model
        print(f"[eval] LLM モデル差し替え: {args.llm_model}", flush=True)

    if args.no_replace:
        cfg["whisper"]["_replacement_rules"] = []
        print("[eval] 決定的置換を無効化", flush=True)

    tr = vi.Transcriber(cfg)
    fm = vi.Formatter(cfg)

    results = []
    asr_hits = 0
    final_hits = 0
    truncations = 0

    for c in cases:
        audio = decode_audio(str(AUDIO_DIR / f"{c['id']}.wav"), sampling_rate=16000)
        raw, whisper_ms = tr.transcribe(audio)

        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            formatted, llm_ms = fm.format(raw, c["prompt_key"])
        guard = next(
            (g for g in ("長さ暴走検出", "中国語混入検出", "英語翻訳検出", "エラー")
             if g in buf.getvalue()),
            "",
        )

        in_raw = c["expect"] in raw
        in_final = c["expect"] in formatted
        asr_hits += in_raw
        final_hits += in_final

        # 忠実度: 整形は「削る・句読点を足す」だけのはずなので、raw から大きく
        # 離れたら整形器が暴走している。固有名詞の有無だけ見ていると
        # 「白照会。」のように文を切り詰めた出力を OK と誤判定するため必須。
        import difflib
        sim = difflib.SequenceMatcher(None, raw, formatted).ratio()
        truncated = len(formatted) < len(raw) * 0.7
        if truncated:
            truncations += 1

        results.append({
            "id": c["id"], "expect": c["expect"], "raw": raw, "formatted": formatted,
            "asr_ok": in_raw, "final_ok": in_final, "guard": guard,
            "similarity": round(sim, 3), "truncated": truncated,
            "whisper_ms": round(whisper_ms), "llm_ms": round(llm_ms),
        })

    total = len(cases)
    mean_sim = round(sum(r["similarity"] for r in results) / total, 3)
    print("\n" + "=" * 84)
    print(f"{'case':<12} {'認識':<5} {'最終':<5} {'類似':<6} 出力")
    print("-" * 84)
    for r in results:
        mark = lambda ok: " OK " if ok else " -- "
        note = f"  [{r['guard']}]" if r["guard"] else ""
        if r["truncated"]:
            note += "  ★文が切り詰められた"
        print(
            f"{r['id']:<12} {mark(r['asr_ok']):<5} {mark(r['final_ok']):<5} "
            f"{r['similarity']:<6.2f} {r['formatted'][:38]}{note}"
        )
    print("-" * 84)
    print(f"認識層(Whisper単体で正解表記): {asr_hits}/{total}")
    print(f"最終(LLM整形まで通した後)    : {final_hits}/{total}")
    print(f"整形層の寄与                 : {final_hits - asr_hits:+d}")
    print(f"忠実度(raw との平均類似度)   : {mean_sim}  ※1.0に近いほど原文を保っている")
    print(f"文の切り詰め                 : {truncations}/{total}  ※0であるべき")

    out = {
        "label": args.label,
        "total": total,
        "asr_score": asr_hits,
        "final_score": final_hits,
        "mean_similarity": mean_sim,
        "truncations": truncations,
        "whisper_model": cfg["whisper"]["model"],
        "llm_model": cfg["llm"]["model"],
        "use_hotwords": cfg["whisper"].get("use_hotwords", False),
        "prompts": args.prompts or "config.json",
        "results": results,
    }
    safe = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in args.label)
    path = RESULTS_DIR / f"{safe}.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n保存: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
