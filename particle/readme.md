# Particle System

ブラウザで動作するリアルタイム3Dパーティクルシミュレーター。  
エミッターの制御・物理演算・レンダリング設定をGUIで調整し、PNG連番をZIPでエクスポートできます。

---

## 起動方法

`particle.html` をブラウザで開くだけで動作します。インストール・サーバー不要。  
推奨ブラウザ：Chrome / Edge（最新版）

---

## プリセット

画面上部のボタンからすぐに呼び出せます。

| ボタン | 内容 |
|--------|------|
| 🔥 Fire | 炎 |
| ✨ Sparks | スパーク |
| ❄ Snow | 雪 |
| 💥 ShockWave | 衝撃波 |
| 🌋 Embers | 火の粉 |

プリセットはベースとして使い、各パラメーターを自由に上書きできます。  
**Save / Load スロット**（スロット名を入力して保存）で独自設定を保持できます。

---

## パラメーター一覧

### EMITTER（エミッター）

**Emitter Type**（エミッタータイプ）

| 値 | 説明 |
|----|------|
| Point | 1点から放出 |
| Box | 直方体の範囲から放出 |
| Sphere | 球形の範囲から放出 |
| ▣ Billboard (PNG) | PNG画像をエミッターマスクとして使用 |

**Burst / Cycle**

| パラメーター | 説明 |
|-------------|------|
| Type | Continuous（連続放出）/ Burst（一括放出） |
| Alpha Threshold | Billboardエミッター使用時、放出するピクセルのアルファ閾値 |
| Interval (s) | Burstの繰り返し間隔（秒） |
| Duration (s) | Burstの持続時間（秒） |

**Emission（放出設定）**

| パラメーター | 説明 |
|-------------|------|
| Particles / sec | 毎秒の放出数（直接入力で最大 5,000,000） |
| Velocity | 初速 |
| Vel Random % | 速度のランダム幅（%） |
| Direction | 放出方向（Up / Down / Outward / Inward / Uniform / Twirl） |
| Spread ° | 放出角度の広がり（度） |
| Pitch ° / Yaw ° | エミッターの向き |
| Twirl Speed | Direction=Twirl 時の螺旋速度 |
| Emitter Size X / Y / Z | Box・Sphereエミッターのサイズ |

---

### PARTICLE（パーティクル）

| パラメーター | 説明 |
|-------------|------|
| Lifespan (s) | 寿命（秒） |
| Life Random % | 寿命のランダム幅 |
| Size | サイズ |
| Size Random % | サイズのランダム幅 |
| Size over Life | 寿命に応じたサイズ変化カーブ |
| Feather % | エッジのぼかし量（Sphere・Star・Dust Blob） |
| Opacity | 全体の不透明度 |
| Flicker Speed | 点滅速度（0=オフ、数値が上がるほど速く点滅） |
| Flicker Rand | 各パーティクルが独立してランダムに消えやすくなる度合い（0〜10）。0=ほぼ消えない、10=多くが消える |
| Spin Amplitude | 回転の振れ幅 |
| Spin Random % | 回転速度のランダム幅 |

**Color over Life（寿命カラー）**

Birth / Mid / End の3点でグラデーションを設定します。  
色見本をクリックしてカラーピッカーを開きます。

**Blend Mode**

| 値 | 説明 |
|----|------|
| Additive (Add) | 加算合成。光・炎・発光エフェクト向き |
| Screen | スクリーン合成。柔らかい発光 |
| Normal | 通常合成 |
| Lighten | 明るい方を残す |

**Particle Shape（パーティクル形状）**

| 形状 | 説明 |
|------|------|
| ● Sphere | 円（デフォルト） |
| ─ Line / Streak | 速度方向に伸びる線。Streak Length・Thickness・Taper で調整 |
| ✦ Star | 星形 |
| ☁ Dust Blob | 有機的な塊 |
| ▣ Billboard (PNG) | PNG画像を貼り付け。Color Tint でカラー乗算 |

---

### PHYSICS / AIR（空気抵抗・風）

| パラメーター | 説明 |
|-------------|------|
| Air Resistance | 速度への抵抗係数（大きいほど早く減速） |
| Air Res Rotation | 回転への抵抗係数 |
| Wind X / Y / Z | ワールド座標系での風力 |

**Turbulence Field（乱流フィールド）**

3Dノイズフィールドでパーティクルをランダムにうねらせます。

| パラメーター | 説明 |
|-------------|------|
| Affect Position | 位置への影響強度 |
| Scale | ノイズのスケール（大きいほどなだらか） |
| Evolution | ノイズの時間変化速度 |
| Octaves | ノイズの重ね数（多いほど細かい） |
| Oct Multiplier | 高周波成分の倍率 |
| Affect Rotation | 回転への影響強度 |

---

### PHYSICS / GRAVITY（重力）

| パラメーター | 説明 |
|-------------|------|
| Gravity | 重力加速度。正=下向き、負=上向き |

---

### PHYSICS / BOUNCE（床バウンス）

| パラメーター | 説明 |
|-------------|------|
| Floor | On / Off |
| Floor Position % | 床の高さ（画面の上端0%〜下端100%） |
| Bounciness % | 反発係数 |
| Friction % | 横方向の摩擦 |
| Bounce Rand % | バウンスのランダム幅 |

---

### RENDERING（レンダリング）

| パラメーター | 説明 |
|-------------|------|
| Max Particles | 同時存在できるパーティクルの上限数 |
| Depth of Field | 奥行きに応じたぼかし強度 |
| Motion Blur | モーションブラー強度（0=オフ、100=最大）。実際に通過した軌跡を記録してトレイルを描画します。Turbulenceでくねった動きもそのまま尾になります |
| Focal Length | 透視投影の焦点距離。大きいほど望遠・小さいほど広角 |

**Displacement Field（ディスプレイスメントフィールド）**

描画座標をノイズで歪める視覚エフェクトです。物理のTurbulenceが「速度に力を加えて実際に曲げる」のに対し、こちらは「レンダリング時に見た目だけを揺らす」ため、軌跡や物理挙動には影響しません。Particularの Turbulence Fields に相当します。

| パラメーター | 説明 |
|-------------|------|
| Strength | 歪みの強さ（0=オフ、300=大きく歪む） |
| Scale | ノイズの粒サイズ。小さいと細かくうねうね、大きいとゆったり波打つ |
| Speed | 歪みパターンの変化速度 |

---

### EXPORT / RENDER

PNG連番をZIP圧縮してダウンロードします。  
エクスポートはリアルタイムと切り離された固定dtシミュレーションで動作するため、PCの処理速度に関係なく一定品質で出力されます。

| パラメーター | 説明 |
|-------------|------|
| Width / Height (px) | 出力解像度。Auto ボタンでアスペクト比を自動計算 |
| FPS | 12 / 24 / 30 / 60 fps から選択 |
| Duration (s) | 書き出す秒数 |
| BG Type | Transparent（アルファ付きPNG）/ Black / Custom color |
| ● REC（リアルタイム録画） | チェックを入れてエフェクトを再生し、■ 保存 で停止するとリアルタイム録画をZIPで保存 |
| ▶ Render & Export ZIP | 設定した解像度・FPS・秒数でオフスクリーンレンダリングしてZIP出力 |

---

## 操作

| 操作 | 動作 |
|------|------|
| キャンバスをクリック | エミッター位置を移動 |
| キャンバスをドラッグ | エミッター位置をリアルタイムに動かす |
| 数値ラベルをクリック | 直接入力モード（Enterで確定、Escでキャンセル） |
| セクションヘッダーをクリック | セクションの開閉 |

---

## 技術仕様

- 純粋なHTML + JavaScript（Canvas 2D API）、依存ライブラリなし
- 3D透視投影（焦点距離可変）
- Catmull-Romスプライン補間によるモーションブラー
- fBmノイズ（フラクタルブラウン運動）による乱流フィールド
- エクスポートは固定dt（デターミニスティック）シミュレーション
