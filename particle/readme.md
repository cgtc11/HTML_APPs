# Particle System

インタラクティブな3Dパーティクルシステム。ブラウザ上で動作するシングルHTMLファイル。

---

## 起動

`particle.html` をブラウザで開くだけで動作します。インストール不要。

---

## プリセット

| プリセット | 説明 |
|---|---|
| 🔥 Fire | 炎。BoxエミッターからUpward方向に発生 |
| 💨 Smoke | 煙。Outward方向に拡散、Sphere type |
| 🌫 Dust | 埃。BoxエミッターにDust Blob type |
| ✨ Sparks | 火花。Line / Streak type |
| ✨ Magic | 魔法のパーティクル。Star type |
| ❄ Snow | 雪。BoxエミッターにDust Blob type |

プリセットは上部ドロップダウンから選択。スロットに保存・呼び出し可能。

---

## パネル構成

### Emitter（エミッター）

| 項目 | 説明 |
|---|---|
| Type | Point / Box / Sphere / **Billboard (PNG)** |
| Emitter PNG | Billboard選択時に表示。PNGの不透明ピクセルからパーティクル発生 |
| Alpha Threshold | 発生対象とするピクセルの最小アルファ値（1〜254）|
| Interval / Duration | バースト発射のサイクル制御 |
| Particles / sec | 毎秒発生数（初期範囲 1〜1,000、最大 50,000） |
| Velocity / Vel Random % | 初速と速度ランダム幅 |
| Direction | Uniform / Upward / Downward / Outward |
| Spread ° | 発射角度の広がり |
| Emitter Size X/Y/Z | エミッター形状のスケール（Billboard選択時は X=100, Y=100, Z=0 で初期化） |

### Particle（パーティクル）

| 項目 | 説明 |
|---|---|
| Lifespan / Life Random % | 寿命と寿命ランダム幅 |
| Size / Size Random % | サイズと個体差 |
| Size over Life | Linear / Ease / Hold / Bell / Grow / Late |
| Feather % | エッジのぼかし量 |
| Opacity | 不透明度 |
| Spin Amplitude / Spin Random % | 回転速度と個体差 |
| Color 1 / 2 / 3 | ライフに沿った3色グラデーション |
| Blend Mode | Source-over / Lighter など |
| **Type** | **● Sphere / ─ Line / ✦ Star / ☁ Dust Blob / ▣ Billboard** |
| Streak Length / Thickness / Taper | Line type のオプション |
| Billboard PNG | Billboard type で使用する画像 |

#### Particle Type 詳細

- **Sphere** — 標準の丸いパーティクル。Featherでグロー表現
- **Line / Streak** — 速度方向に伸びるストリーク。テーパー有無切替可
- **Star (✦)** — 4本スパイクの星形。パーティクルごとに縦横比・中心サイズをランダム変化
- **Dust Blob (☁)** — 3つの円が重なるクラスター形状。クラスターの向きがパーティクルごとにランダム変化
- **Billboard (PNG)** — PNGをそのまま描画。カラーオーバーライフでティント可

### Physics（物理）

| 項目 | 説明 |
|---|---|
| Air Resistance | 空気抵抗（初期 0〜100、最大 500） |
| Air Res Random % | 空気抵抗の個体差 |
| Wind X/Y/Z | 風力（初期 ±50、最大 ±500） |
| Turbulence AP / Scale / Evolution / Octaves | ノイズ乱流の設定 |
| Gravity | 重力（負=上昇） |
| Floor / Bounce / Friction / Bounce Random | 床との衝突設定 |
| Max Particles | 最大同時パーティクル数 |
| Depth of Field | ボケ効果（奥行きに応じてサイズ・透明度変化） |
| Focal Length | 遠近感のスケール |

---

## スライダーの直接入力（ダブルクリック）

数値ラベルをダブルクリックするとテキスト入力モードになります。  
入力した値がそのままスライダーの **新しいMAX（またはMIN）** になります。

- 正の値を入力 → MAX を更新（縮小・拡大どちらも対応）
- 負の値を入力 → MIN を更新（Wind XYZ など）
- `Enter` で確定 / `Escape` でキャンセル

各スライダーのハード上限：

| スライダー | 初期範囲 | ハード上限 |
|---|---|---|
| Particles / sec | 1〜1,000 | 50,000 |
| Emitter Size XYZ | 0〜500 | 5,000 |
| Air Resistance | 0〜100 | 500 |
| Wind XYZ | ±50 | ±500 |

---

## プリセットの保存・読み込み

- **スロット（1〜5）** に名前付きで保存可能
- 右上メニューから **JSON エクスポート / インポート** に対応
- JSONファイルは複数スロットをまとめて保存

---

## 録画・エクスポート

- **REC ボタン** でフレームシーケンス（PNG連番）をZIPとして書き出し
- フレームレート・解像度・時間を指定可能

---

## 動作環境

モダンブラウザ（Chrome / Firefox / Safari / Edge）。外部ライブラリ不要。
