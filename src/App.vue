<script setup>
import { computed, ref, watch } from 'vue';
import { layout, buildDocx, buildPdf } from './lib/genkou.js';
import GenkouSheet from './components/GenkouSheet.vue';

const STORE = 'genkou-src-v1';
const FONTS = ['ＭＳ 明朝', '游明朝', 'ＭＳ ゴシック', '游ゴシック'];
const FORMATS = {
  docx: { build: buildDocx, type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
  pdf: { build: buildPdf, type: 'application/pdf' },
};
const PLACEHOLDER = `# 第1問（設問1） 100字
強みは、①長年培った高い技術力、②地域顧客との信頼関係。弱みは、①後継者不在、②営業力の不足である。

# 第1問（設問2） 80字
...`;

function load() {
  try {
    return localStorage.getItem(STORE) || '';
  } catch {
    return '';
  }
}

const src = ref(load());
const title = ref('');
const limit = ref(null);
const font = ref(FONTS[0]);
const brackets = ref(false);
const keepNewlines = ref(false);
const srcEl = ref(null);
// 設問ごとの「書き出し1マスあけ」（設問の番号 → true）
const INDENT_STORE = 'genkou-indents-v1';
function loadIndents() {
  try {
    return JSON.parse(localStorage.getItem(INDENT_STORE)) || {};
  } catch {
    return {};
  }
}
const indents = ref(loadIndents());
function setIndent(i, on) {
  indents.value = { ...indents.value, [i]: on };
  try {
    localStorage.setItem(INDENT_STORE, JSON.stringify(indents.value));
  } catch {
    /* 保存できない環境では何もしない */
  }
}

watch(src, (v) => {
  try {
    localStorage.setItem(STORE, v);
  } catch {
    /* 保存できない環境では何もしない */
  }
});

const blocks = computed(() =>
  layout(src.value, {
    brackets: brackets.value,
    keepNewlines: keepNewlines.value,
    indents: indents.value,
    defaultLimit: limit.value > 0 ? limit.value : null,
  }),
);

function download(ext) {
  if (!blocks.value.length) {
    srcEl.value?.focus();
    return;
  }
  const t = title.value.trim();
  const bytes = FORMATS[ext].build(blocks.value, { title: t, font: font.value });
  const url = URL.createObjectURL(new Blob([bytes], { type: FORMATS[ext].type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = (t || '再現答案').replace(/[\\/:*?"<>|]/g, '_') + '.' + ext;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
</script>

<template>
  <header>
    <h1>中小企業診断士2次試験解答用紙メーカー</h1>
    <p>中小企業診断士 2次試験の再現答案を貼り付けると、1行20マスの解答用紙に1マスずつ書いた形で Word（.docx）・PDF に出力します。</p>
  </header>
  <main>
    <section class="panel">
      <label for="src">再現答案</label>
      <textarea id="src" ref="srcEl" v-model="src" :placeholder="PLACEHOLDER"></textarea>
      <div class="row">
        <label>タイトル <input v-model="title" type="text" placeholder="例：令和7年度 事例Ⅰ 再現答案" /></label>
      </div>
      <div class="row">
        <label>見出しなし時の字数 <input v-model.number="limit" type="number" min="0" step="10" placeholder="自動" /></label>
        <label>フォント
          <select v-model="font">
            <option v-for="f in FONTS" :key="f" :value="f">{{ f }}</option>
          </select>
        </label>
      </div>
      <div class="row">
        <label><input v-model="brackets" type="checkbox" /> 閉じ括弧（」）など）も行頭に置かない</label>
        <label><input v-model="keepNewlines" type="checkbox" /> 改行を残す（改行で次の行へ）</label>
      </div>
      <div class="row">
        <button @click="download('docx')">Word でダウンロード</button>
        <button @click="download('pdf')">PDF でダウンロード</button>
      </div>
      <div class="help">
        ・全角1文字＝1マス。英数字など半角は1マスに2文字入ります（例：<code>RO</code><code>E1</code><code>0%</code>）。<br />
        ・行頭に来る句読点（、。）は前の行の最後のマスに一緒に入れます。<br />
        ・<code>#</code> で始まる行は設問の見出しです。見出しに「100字」のように書くとその字数分のマス目を用意し、超過を赤で示します。<br />
        ・プレビューの各設問のチェックボックスで、その設問だけ段落の書き出しを1マスあけられます（「改行を残す」なら各段落、それ以外は冒頭）。<br />
        ・見出しがない場合は全体を1つの答案として扱います。貼り付け時の改行は既定で詰めます。<br />
        ・PDF は明朝＝平成明朝、ゴシック＝平成角ゴシック（閲覧環境のフォントで表示）になります。<br />
        ・入力内容はブラウザの中だけで処理され、どこにも送信されません。
      </div>
    </section>
    <section class="panel">
      <div class="preview">
        <GenkouSheet v-for="(b, i) in blocks" :key="i" :block="b" @update:indent="setIndent(i, $event)" />
        <p v-if="!blocks.length" class="empty">左に答案を貼り付けるとここにプレビューが出ます。</p>
      </div>
    </section>
  </main>
</template>
