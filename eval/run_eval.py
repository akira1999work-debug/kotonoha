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
    args = ap.parse_args()

    RESULTS_DIR.mkdir(parents=True, exist_ok=True)

    if args.compare:
        rows = []
        for f in sorted(RESULTS_DIR.glob("*.json")):
            d = json.loads(f.read_text(encoding="utf-8"))
            rows.append((d["label"], d["asr_score"], d["final_score"], d["total"]))
        if not rows:
            print("保存済み結果なし")
            return 0
        print(f"{'施策':<28} {'認識層':>8} {'最終':>8}")
        print("-" * 48)
        for label, asr, fin, total in rows:
            print(f"{label:<28} {asr:>4}/{total:<3} {fin:>4}/{total:<3}")
        return 0

    cases = json.loads((EVAL_DIR / "cases.json").read_text(encoding="utf-8"))["cases"]
    ensure_audio(cases)

    import voice_input as vi
    from faster_whisper.audio import decode_audio

    cfg = vi.load_config()
    tr = vi.Transcriber(cfg)
    fm = vi.Formatter(cfg)

    results = []
    asr_hits = 0
    final_hits = 0

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

        results.append({
            "id": c["id"], "expect": c["expect"], "raw": raw, "formatted": formatted,
            "asr_ok": in_raw, "final_ok": in_final, "guard": guard,
            "whisper_ms": round(whisper_ms), "llm_ms": round(llm_ms),
        })

    total = len(cases)
    print("\n" + "=" * 78)
    print(f"{'case':<12} {'認識':<5} {'最終':<5} 出力")
    print("-" * 78)
    for r in results:
        mark = lambda ok: " OK " if ok else " -- "
        note = f"  [{r['guard']}]" if r["guard"] else ""
        print(f"{r['id']:<12} {mark(r['asr_ok']):<5} {mark(r['final_ok']):<5} {r['formatted'][:44]}{note}")
    print("-" * 78)
    print(f"認識層(Whisper単体で正解表記): {asr_hits}/{total}")
    print(f"最終(LLM整形まで通した後)    : {final_hits}/{total}")
    print(f"整形層の寄与                 : {final_hits - asr_hits:+d}")

    out = {
        "label": args.label,
        "total": total,
        "asr_score": asr_hits,
        "final_score": final_hits,
        "whisper_model": cfg["whisper"]["model"],
        "llm_model": cfg["llm"]["model"],
        "use_hotwords": cfg["whisper"].get("use_hotwords", False),
        "results": results,
    }
    safe = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in args.label)
    path = RESULTS_DIR / f"{safe}.json"
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n保存: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
