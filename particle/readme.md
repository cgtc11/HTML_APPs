# 🔥 Particle Simulator — Particular Style

ブラウザで動く、After Effects「Particular」風の3Dパーティクルシミュレーターです。  
インストール不要。HTMLファイルをダブルクリックするだけで動きます。

![Particle Simulator](https://img.shields.io/badge/動作環境-ブラウザ-blue) ![License](https://img.shields.io/badge/license-MIT-green)

---

## ✨ 特徴

- **6種類のプリセット** — Fire / Smoke / Dust / Sparks / Magic / Snow
- **リアルタイム3Dシミュレーション** — 視差・被写界深度・フォーカスレングスに対応
- **豊富なパラメータ** — エミッター、パーティクル、物理、レンダリングを細かく制御
- **オフラインレンダリング** — PNG連番をZIPでエクスポート（透過アルファ対応）
- **リアルタイム録画（REC）** — 再生中の映像をWebMで録画・保存
- **プリセット保存／読み込み** — 設定をJSONで書き出し・読み込み可能
- **インストール不要** — 単一HTMLファイル、外部サーバー不要

---

## 🚀 使い方

### 起動

```
particle_sim_particular.html をブラウザで開く
```

Chrome / Edge / Firefox 最新版を推奨。Safari は一部機能が動作しない場合があります。

### 基本操作

| 操作 | 内容 |
|---|---|
| クリック | エミッター位置を移動 |
| ドラッグ | エミッターを動かしながら放出 |
| 右パネル各スライダー | パラメータをリアルタイムで調整 |
| プリセットボタン | シーンをリセットしてプリセットを適用 |

---

## 🎛️ パラメータ解説

### Emitter（エミッター）

| パラメータ | 内容 |
|---|---|
| Type | エミッターの形状（Point / Box / Sphere） |
| Interval (s) | バースト間隔。0 で連続放出 |
| Duration (s) | 1バーストの放出時間（最小 0.01s） |
| Particles / sec | 毎秒の放出数 |
| Velocity | 初速 |
| Vel Random % | 初速のランダム幅 |
| Direction | 放出方向（Uniform / Upward / Downward / Outward） |
| Spread ° | 放出角度の広がり |
| Emitter Size X/Y/Z | エミッター領域のサイズ |

### Particle（パーティクル）

| パラメータ | 内容 |
|---|---|
| Lifespan (s) | 寿命 |
| Life Random % | 寿命のランダム幅 |
| Size | 初期サイズ |
| Size Random % | サイズのランダム幅 |
| Size over Life | 寿命に沿ったサイズ変化カーブ |
| Feather % | エッジのぼかし具合 |
| Opacity | 不透明度 |
| Spin Amplitude | スピン速度 |
| Color over Life | Birth / Mid / End の3点カラー補間 |
| Blend Mode | 合成モード（Additive / Normal / Screen） |
| Type | 形状（Sphere / Line・Streak / Billboard PNG） |

### Physics / Air（空気・風）

| パラメータ | 内容 |
|---|---|
| Air Resistance | 空気抵抗（速度の減衰） |
| Wind X / Y / Z | 風の強さ・方向 |
| Turbulence | fBmノイズによる乱流フィールド |

### Physics / Gravity（重力）

| パラメータ | 内容 |
|---|---|
| Gravity | 重力加速度（負値で上昇） |

### Physics / Bounce（バウンス）

| パラメータ | 内容 |
|---|---|
| Floor | 床の有効・無効 |
| Floor Position % | 床のY位置 |
| Bounciness % | 反発係数 |
| Friction % | 床との摩擦 |

### Rendering（レンダリング）

| パラメータ | 内容 |
|---|---|
| Max Particles | 同時存在できる最大パーティクル数 |
| Depth of Field | 被写界深度の強さ |
| Focal Length | 焦点距離（透視投影） |

---

## 📤 Export / Render

### オフラインレンダリング（ZIP出力）

1. **Export / Render** セクションを開く
2. 解像度・FPS・秒数・背景色を設定
3. **▶ Render & Export ZIP** をクリック
4. 全フレームをPNGで描画後、ZIPを自動ダウンロード

> 透過背景（アルファ）を選択すると、After Effects や DaVinci Resolve に直接読み込めます。

### リアルタイム録画（REC）

1. **Export / Render** セクション内の **● REC リアルタイム録画** をオンにする
2. シミュレーションが録画開始（フレームカウンターが表示される）
3. **■ 保存** を押すと WebM 動画をダウンロード

---

## 💾 プリセットの保存・読み込み

| ボタン | 内容 |
|---|---|
| 💾 Save | 現在の設定に名前を付けてブラウザに保存 |
| 📂 Load | 保存済みプリセットを一覧表示して適用 |
| ⬇ Export | 設定をJSONファイルとして書き出し |
| ⬆ Import | JSONファイルから設定を読み込み |

---

## 🗂️ ファイル構成

```
particle_sim_particular.html   # メインファイル（これ1つで完結）
README.md
```

外部依存は [JSZip](https://stuk.github.io/jszip/)（CDN読み込み）のみです。

---

## 🛠️ 動作環境

| ブラウザ | 対応状況 |
|---|---|
| Chrome 100+ | ✅ 推奨 |
| Edge 100+ | ✅ |
| Firefox 100+ | ✅ |
| Safari | ⚠️ REC機能が動作しない場合あり |

---

## 📄 ライセンス

MIT License
