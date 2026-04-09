# Kotonoha（言の葉）

ローカル完結のリアルタイム日本語音声入力ツール。
Right Alt 2 回で画面下にバーを出し、1 回押すごとに録音開始／停止。
文字起こしは [faster-whisper](https://github.com/SYSTRAN/faster-whisper)、
整形は [Ollama](https://ollama.com/) 上のローカル LLM（既定: Qwen2.5 7B）。
ネット接続なし、API キーなし、すべて自分の PC で完結する。

## 特長

- **常駐オーバーレイ**: 画面下に薄いピル型のバー。録音中は波形が動く
- **ホットキー**: Right Alt ダブルタップで表示切替、シングルタップで録音トグル
- **リアルタイム文字起こし**: 録音中でも途中経過を逐次表示
- **LLM 整形**: フィラー除去と句読点補完だけの「最小整形」プロンプトで、元の言い回しを壊さない
- **アプリ別プロンプト切替**: Claude Code / VSCode / Cursor / LINE / Discord などアクティブウィンドウで整形ルールを変える
- **ユーザー辞書**: 固有名詞の誤認識を `user_dictionary.json` に登録していける
- **自動貼り付け**: 文字起こし結果をクリップボード＆Ctrl+V で即座に入力欄へ
- **システムトレイ常駐**

## 動作環境

- **OS**: Windows 10 / 11（動作確認は Windows 11）
- **Python**: 3.10 以降
- **GPU**: NVIDIA GPU 推奨（CUDA）。CPU でも動くが Whisper の速度は落ちる
- **メモリ**: Whisper large-v3-turbo + Qwen 2.5 7B を同時に動かすと VRAM 約 8 GB 程度。足りない場合は下記「軽量設定」参照

## インストール

### 1. Python 依存

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> `pycairo` のインストールで失敗する場合は Windows では
> [こちらのビルド済み wheel](https://www.lfd.uci.edu/~gohlke/pythonlibs/#pycairo)
> から `.whl` を落として `pip install pycairo-xxx.whl` で入れる。

### 2. Ollama をインストール

https://ollama.com/download から Windows 版をインストールして、モデルを pull:

```powershell
ollama pull qwen2.5:7b-instruct-q4_K_M
```

> 軽くしたい場合は `qwen2.5:3b-instruct-q4_K_M` などに差し替えて、
> `config.json` の `llm.model` も同じ名前に書き換える。

### 3. Whisper モデル

初回起動時に `faster-whisper` が自動でダウンロードする。既定は `large-v3-turbo`。
VRAM が足りない場合は `config.json` の `whisper.model` を `medium` や `small` に変える。

### 4. CUDA が無い場合

`config.json` の `whisper` セクションを以下に変更:

```json
"whisper": {
  "model": "medium",
  "device": "cpu",
  "compute_type": "int8",
  ...
}
```

## 起動

```powershell
python voice_input.py
```

コンソールなしで起動したい場合は `pythonw.exe voice_input.py`。

### ショートカット作成 / スタートアップ登録

付属の `create_shortcut.ps1` でショートカットを作れる。デフォルトはデスクトップに置く、**opt-in でスタートアップフォルダに登録**できる。

```powershell
# デスクトップにショートカットを作成 (デフォルト)
.\create_shortcut.ps1

# Windows のスタートアップフォルダに登録 (ログイン時に自動起動)
.\create_shortcut.ps1 -Startup
```

`-Startup` を付けると Windows ログイン時に自動で常駐が立ち上がり、Whisper と Ollama のモデルロードが**バックグラウンドで完了してから**初回の録音を開始できるので、起動待ちの体感がゼロになる。

解除するには `explorer shell:startup` で開いたフォルダから `VoiceInput.lnk` を削除する。

## 使い方

| 操作 | 動作 |
|------|------|
| Right Alt × 2（ダブルタップ） | バーの表示／非表示 |
| Right Alt × 1（バー表示中） | 録音開始／停止 |
| Right Alt + Ctrl | 終了 |
| バー上のマイクアイコン | 録音開始／停止 |
| バー上の × | バーを隠す（常駐は継続） |
| システムトレイ | 右クリックメニューから操作 |

録音停止後、Whisper で文字起こし → LLM で整形 → クリップボードにコピー → アクティブウィンドウに Ctrl+V で自動貼り付けされる。

## 設定

すべて `config.json` に集約されている。主な項目:

| キー | 説明 |
|------|------|
| `hotkey` | ホットキー（既定: `right alt`） |
| `max_duration_sec` | 1 回の録音の最大秒数 |
| `realtime.enabled` | リアルタイム文字起こし ON/OFF |
| `whisper.model` | Whisper モデル名 |
| `whisper.device` | `cuda` or `cpu` |
| `llm.model` | Ollama のモデル名 |
| `llm.keep_alive` | Ollama にモデルを載せっぱなしにする時間 |
| `prompts` | 整形プロンプト（default / code / casual） |
| `app_routing` | アクティブアプリ別にどのプロンプトを使うか |
| `output.auto_paste` | Ctrl+V 自動送信 ON/OFF |

### アプリ別プロンプト

`app_routing` でアクティブウィンドウの実行ファイル名をキーにして、使うプロンプトを切り替えられる。例:

```json
"app_routing": {
  "Code.exe": "code",
  "Cursor.exe": "code",
  "LINE.exe": "casual",
  "Discord.exe": "casual"
}
```

「コード書いてる時は敬語にしない、LINE では少しだけ丁寧に」みたいな微妙な出し分けができる。

## ユーザー辞書

`user_dictionary.json` に固有名詞を登録しておくと、Whisper の `initial_prompt` と LLM の補正ヒントとして使われる。

```json
{
  "term": "Next.js",
  "readings": ["ネクストジェーエス", "ネクストジェイエス"],
  "category": "tool",
  "priority": 7
}
```

誤認識が出たら `readings` に観測された誤りパターンを追記していけばよい。

## トラブルシューティング

- **Ctrl+V で貼り付けされない** → IME がオフになっているか、アクティブウィンドウに入力フォーカスが無い可能性。`output.auto_paste` を `false` にしてクリップボード貼り付けで運用する
- **起動が遅い** → 初回は Whisper モデルのダウンロードが走る。2 回目以降でも Whisper + Ollama の cold load で数秒かかるので、起動待ちを消したい場合は `.\create_shortcut.ps1 -Startup` で Windows スタートアップに登録してログイン時に裏でロードさせる
- **VRAM 不足で落ちる** → `whisper.model` を `medium` → `small` に、`llm.model` を `qwen2.5:3b-instruct-q4_K_M` に落とす
- **Ollama に繋がらない** → Ollama が起動しているか確認（タスクトレイにアイコンが出る）

## ライセンス

MIT License. 詳細は [LICENSE](./LICENSE)。
