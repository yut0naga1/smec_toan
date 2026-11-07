# 中小企業診断士2次試験解答用紙メーカー

中小企業診断士2次試験の再現答案を貼り付けると、1行20マスの解答用紙に1マスずつ書いた形で Word（.docx）・PDF を出力するサイト。Vue 3 + Vite。入力はブラウザ内で処理し、外部に送信しない。

公開URL: https://yut0naga1.github.io/smec_toan/

## 仕様

- 全角1文字＝1マス。半角英数字は1マスに2文字。
- 行頭に来る句読点（、。，．）は前の行の20マス目に同居させる。閉じ括弧も同様に扱うオプションあり。
- `#` で始まる行は設問見出し。`# 第1問 100字` のように字数を書くと、その字数分のマス目を用意し超過を警告する。
- 100字ごとに右端へ目盛り（100, 200, …）を付ける。

## 開発

```sh
npm install
npm run dev      # 開発サーバー
npm test         # ロジックのテスト（node --test）
npm run build    # dist/ に出力
```

`main` に push すると GitHub Actions（`.github/workflows/deploy.yml`）がテスト・ビルドして Pages に公開する。

## 構成

- `src/lib/genkou.js` — マス割り付けと .docx / .pdf 生成（画面から独立、依存なし）
  - docx: 無圧縮ZIPを自前で組み立て
  - pdf: 非埋め込みの標準日本語フォント（HeiseiMin-W3 / HeiseiKakuGo-W5, UniJIS-UCS2-H）。1文字ずつ座標指定するため代替フォントでもマスがずれない
- `src/App.vue` — 入力フォームとダウンロード
- `src/components/GenkouSheet.vue` — 1設問分の解答用紙プレビュー
- `test/genkou.test.mjs` — ロジックのテスト
