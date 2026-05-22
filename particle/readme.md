# Particle Simulator — Particular Style

Trapcode Particular にインスパイアされた、ブラウザで動作するリアルタイム3Dパーティクルシミュレーター。  
単一HTMLファイルで完結し、インストール不要。ローカルで開くだけで使えます。

---

## 起動方法

```
particle_sim_particular.html をブラウザで開く
```

依存ライブラリなし。インターネット接続不要。Chrome / Edge / Firefox / Safari 対応。

---

## 基本操作

| 操作 | 内容 |
|------|------|
| キャンバス クリック / ドラッグ | エミッター位置を移動（既存パーティクルはその場に残る） |
| 数値ラベルをダブルクリック | 直接数値入力（スライダー範囲外の値も入力可能） |

---

## プリセット

パネル上部のボタンで6種類のプリセットを切り替えられます。  
切り替えると既存パーティクルはクリアされ、全パラメーターがリセットされます。

| プリセット | 説明 |
|-----------|------|
| 🔥 Fire | 揺らめく炎 |
| 💨 Smoke | ふわふわと広がる煙 |
| 🌫 Dust | 乱流に漂う砂埃 |
| ✨ Sparks | 重力で落下・床バウンドする火花 |
| 🌀 Magic | 虹色に輝く魔法エフェクト |
| ❄ Snow | 横風に揺れながら降る雪 |

---

## パラメーター詳細

### Burst / Cycle
パーティクルの発射タイミングを制御します。

| パラメーター | 範囲 | 説明 |
|------------|------|------|
| Interval (s) | 0〜30 | 発射間隔（0 = 無限発射） |
| Duration (s) | 0.05〜10 | 1サイクルあたりの発射時間 |

**動作例：** Interval=5, Duration=1 → 1秒発射 → 5秒停止 → 繰り返し  
**爆発用途：** Duration=0.1, Particles/sec=50000 で一瞬大量放出

---

### Emitter（エミッター）

| パラメーター | 説明 |
|------------|------|
| Type | Point / Box / Sphere — エミッターの形状 |
| Particles / sec | 1秒あたりの発生数（最大50,000） |
| Velocity | 初速度 |
| Vel Random % | 初速度のランダム幅 |
| Direction | Uniform（全方向）/ Upward / Downward / Outward |
| Spread ° | 発射角度の広がり（0〜180°） |
| Emitter Size X/Y/Z | エミッター形状のサイズ（Box/Sphere時に有効） |

---

### Particle（パーティクル）

| パラメーター | 説明 |
|------------|------|
| Lifespan (s) | 寿命（秒） |
| Life Random % | 寿命のランダム幅 |
| Size | 基本サイズ |
| Size Random % | サイズのランダム幅 |
| Size over Life | Linear↓ / Ease out / Hold→Drop / Bell∩ / Grow→Drop / Late Drop |
| Feather % | エッジのソフトさ（高いほどぼんやり） |
| Opacity | 不透明度 |
| Spin Amplitude | 回転速度 |
| Spin Random % | 回転方向・速度のランダム幅 |
| Color Birth/Mid/End | 寿命に沿った3点カラー補間 |
| Blend Mode | Additive（加算）/ Normal / Screen |

#### Particle Shape（パーティクル形状）

| タイプ | 説明 |
|--------|------|
| ● Sphere | ソフトな球体（デフォルト）。Feather でエッジを調整 |
| ─ Line / Streak | 速度ベクトル方向に伸びるストリーク。Streak Length・Thickness・Taper を設定 |
| ▣ Billboard (PNG) | PNG/WebP/GIF 画像をビルボード表示。透過チャンネルそのまま有効。Color Tint ON で Color over Life を乗算適用 |

---

### Physics / Air（空気・乱流）

| パラメーター | 説明 |
|------------|------|
| Air Resistance | 速度への指数減衰ドラッグ（高いほど早く止まる） |
| Air Res Rotation | 回転への個別ドラッグ |
| Wind X / Y / Z | 定常風の方向・強さ |
| **Affect Position** | 乱流フィールドが位置に与える強さ（0=無効） |
| Scale | 乱流の空間スケール（大きいほど大きな渦） |
| Evolution | 乱流フィールドの時間変化速度 |
| Octaves | fBm のオクターブ数（多いほど細かい乱流、重くなる） |
| Oct Multiplier | 各オクターブの振幅減衰率 |
| Affect Rotation | 乱流フィールドが回転に与える強さ |

> 乱流は3次元 fBm（フラクタルブラウン運動）ノイズによって実装されており、空間的に連続した有機的な揺らぎを生成します。

---

### Physics / Gravity

| パラメーター | 説明 |
|------------|------|
| Gravity | 重力の強さと方向（負 = 上昇、正 = 落下、0 = 無重力） |

---

### Physics / Bounce

| パラメーター | 説明 |
|------------|------|
| Floor | 床の ON/OFF |
| Floor Position % | 画面に対する床の高さ（%） |
| Bounciness % | 跳ね返り係数 |
| Friction % | 床との摩擦（水平速度の減衰） |
| Bounce Rand % | バウンス係数のランダム幅 |

---

### Rendering

| パラメーター | 説明 |
|------------|------|
| Max Particles | 同時存在できる最大パーティクル数（最大500,000） |
| Depth of Field | 被写界深度効果（Z方向の距離でサイズ拡大・透明度低下） |
| Focal Length | 3Dパースペクティブの焦点距離（大きいほど望遠） |

---

## 設定の保存・読み込み

| ボタン | 機能 |
|--------|------|
| ↺ Reset | 現在のプリセットに戻す（パーティクルもクリア） |
| 💾 Save | 名前をつけてブラウザの localStorage に保存 |
| 📂 Load | 保存済みリストから読み込み・削除 |
| ⬇ Export | 全パラメーターを JSON ファイルとしてダウンロード |

> **注意：** 保存データはブラウザの localStorage に保存されます。  
> 異なるブラウザや別のPCへの移行は Export（JSON）を使用してください。  
> 画像（Billboard）の情報は設定には含まれません。再読み込みが必要です。

---

## パフォーマンスの目安

| パーティクル数 | 目安 |
|--------------|------|
| 〜5,000 | 快適（ほぼすべての環境） |
| 5,000〜30,000 | 普通のPC・スマートフォンで動作 |
| 30,000〜100,000 | 高性能PC推奨。Sphere タイプが最軽量 |
| 100,000〜 | GPU性能次第。処理落ち覚悟の実験領域 |

**軽量化のヒント：**
- Turbulence の Octaves を下げる（最も効果的）
- Sphere タイプを使う（Billboard が最重量）
- Feather % を下げる（描画レイヤーが減る）

---

## 技術仕様

- **レンダリング：** HTML5 Canvas 2D API
- **3D：** パースペクティブ投影（焦点距離可変）、Zソート（ペインターズアルゴリズム）
- **乱流：** 3次元 Value Noise による fBm（フラクタルブラウン運動）
- **物理：** 指数減衰ドラッグ、定常風、乱流ベクトルフィールド、重力、床バウンス
- **保存：** localStorage（ブラウザ内）+ JSON エクスポート
- **依存：** なし（単一 HTML ファイル）

