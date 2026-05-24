# 🎇 Particle Designer

リアルタイム3Dパーティクルシミュレーター。ブラウザで動作する単一HTMLファイル。

---

## 起動方法

`particle.html` をブラウザで開くだけで動作します。インストール・サーバー不要。

---

## 基本操作

| 操作 | 内容 |
|---|---|
| 左クリック / ドラッグ | エミッターを移動 |
| ⛶ Full ボタン | フルスクリーン表示（パーティクルのみ） |
| 右クリック / ESC | フルスクリーン終了 |
| 数値ラベルをダブルクリック | 直接数値入力（拡張範囲まで入力可能） |

---

## プリセット

画面上部のボタンで切り替え。

| ボタン | 内容 |
|---|---|
| 🔥 Fire | 炎 |
| ✨ Sparks | 火花 |
| ❄ Snow | 雪 |
| 💥 ShockWave | 衝撃波（Outward方向のリング状爆発） |
| 🌋 Embers | 火の粉（Twirl方向でゆらぎながら舞い上がる） |

### プリセットの保存・読み込み

- **⬇ Export** — 現在のスロット一覧を JSON ファイルとして書き出し
- **⬆ Import** — JSON ファイルを読み込み、スロットに追加
- スロットリストのプリセット名をクリックで呼び出し、ダブルクリックで名前変更、× で削除

---

## パラメーター解説

### Emitter（エミッター）

#### Burst / Cycle
| パラメーター | 説明 |
|---|---|
| Interval (s) | バースト間隔（秒）。0 で連続発生。ダブルクリックで最大 600 秒まで入力可 |
| Duration (s) | 1バーストあたりの発生時間 |

#### Emission（放出）
| パラメーター | 説明 |
|---|---|
| Particles / sec | 毎秒の発生数。ダブルクリックで最大 50,000 まで入力可 |
| Velocity | 発生速度 |
| Vel Random % | 速度のランダムばらつき |
| Direction | 発射方向（下記参照） |
| Spread ° | 発射角度の広がり（0 = 一方向、180 = 全方向） |
| Pitch ° | 発射方向の前後傾き（−90〜+90、ダブルクリックで±180まで） |
| Yaw ° | 発射方向の水平回転（−90〜+90、ダブルクリックで±180まで） |
| Twirl Speed | Twirl方向選択時の螺旋回転速度 |
| Emitter Size X / Y / Z | エミッターの3次元サイズ（box / sphere タイプ時有効） |

**Direction オプション**

| 値 | 説明 |
|---|---|
| Uniform | 全方向にランダムで飛び散る |
| Upward | 上方向（Pitch / Yaw で傾き調整可能） |
| Downward | 下方向（Pitch / Yaw で傾き調整可能） |
| Outward | エミッター中心から外側に円状に広がる（Shockwave向き） |
| Twirl | 螺旋を描きながら進む。Twirl Speed で速度調整 |

> **Shockwave の作り方**  
> Direction → `Outward`、Spread ° → `0`、Pitch ° で平面の傾きを調整。

---

### Particle（パーティクル）

| パラメーター | 説明 |
|---|---|
| Lifespan (s) | 寿命（秒） |
| Life Random % | 寿命のランダムばらつき |
| Size | パーティクルサイズ |
| Size Random % | サイズのランダムばらつき |
| Size over Life | 寿命中のサイズ変化カーブ（下記参照） |
| Feather % | エッジのぼかし量 |
| Opacity | 不透明度 |
| Spin Amplitude | 回転速度 |
| Spin Random % | 回転方向のランダムばらつき |
| Birth / Mid / End | 誕生・中間・消滅時のカラー（グラデーション） |
| Blend Mode | 描画モード（Additive / Normal / Screen） |

**Size over Life オプション**

| 値 | 説明 |
|---|---|
| Linear ↓ | 直線的に縮小 |
| Ease out | 素早く縮小して消える |
| Hold → Drop | 一定サイズを維持して最後に消える |
| Bell ∩ | 山なりに拡大してから縮小 |
| Grow → Drop | 拡大してから急に消える |
| Late Drop | 最後まで大きいまま急に消える |

**Particle Shape**

| 値 | 説明 |
|---|---|
| ● Sphere | 円形（標準） |
| ─ Line / Streak | 速度方向に伸びるライン（Streak Length / Thickness / Taper で調整） |
| ✦ Star | 星形 |
| ☁ Dust Blob | 不規則な雲状ブロブ（煙・埃向き） |
| ▣ Billboard (PNG) | 画像をパーティクルとして使用（PNG/WebP/GIF対応） |

---

### Physics / Air（空気・乱流）

| パラメーター | 説明 |
|---|---|
| Air Resistance | 空気抵抗（値が大きいほど早く減速） |
| Air Res Rotation | 回転への空気抵抗 |
| Wind X / Y / Z | 風の強さ（3軸） |

**Turbulence Field（乱流フィールド）**

| パラメーター | 説明 |
|---|---|
| Affect Position | 乱流の位置への影響強度 |
| Scale | 乱流の空間スケール（小さいほど細かくゆらぐ） |
| Evolution | 乱流の時間変化速度 |
| Octaves | 乱流の複雑さ（重ね合わせ回数） |
| Oct Multiplier | 各オクターブの振幅倍率 |
| Affect Rotation | 乱流の回転への影響強度 |

---

### Physics / Gravity（重力）

| パラメーター | 説明 |
|---|---|
| Gravity | 重力（正 = 下向き、負 = 上向き浮力） |

---

### Physics / Bounce（バウンス・床）

| パラメーター | 説明 |
|---|---|
| Floor | 床の有効 / 無効 |
| Floor Position % | 床の縦位置（画面上端を0%、下端を100%） |
| Bounciness % | 反発係数 |
| Friction % | 摩擦 |
| Bounce Rand % | 反発のランダムばらつき |

---

### Rendering（レンダリング）

| パラメーター | 説明 |
|---|---|
| Max Particles | 同時存在できるパーティクルの上限数 |
| Depth of Field | ぼかし量（奥行き感） |
| Focal Length | 焦点距離（透視変換の強さ） |

---

### Export / Render（書き出し）

| 項目 | 説明 |
|---|---|
| Width / Height (px) | 出力解像度。Auto ボタンでアスペクト比から自動計算 |
| FPS | フレームレート（12 / 24 / 30 / 60） |
| Duration (s) | 書き出し時間 |
| BG Type | 背景（透過 / 黒 / カスタムカラー） |
| ▶ Render & Export ZIP | 全フレームをPNG連番でZIP書き出し |
| ● REC | リアルタイム録画（WebM形式で保存） |

---

## JSONプリセット形式

Import / Export するJSONの構造：

```json
{
  "version": 1,
  "slots": [
    {
      "name": "🔥 Fire",
      "preset": "fire",
      "data": {
        "pps": 313,
        "vel": 4.8,
        "dir": "up",
        "c1": "#ffee80",
        "c2": "#ff3300",
        "c3": "#220000"
      }
    }
  ]
}
```

`preset` には `fire` / `sparks` / `snow` / `shockwave` / `embers` / `smoke` / `dust` / `magic` が指定可能（UIの初期レイアウトに影響）。

---

## ファイル構成

```
particle.html       メインファイル（これ1つで動作）
README.md           本ドキュメント
presets_all.json    プリセット集（Import で読み込み可能）
```
