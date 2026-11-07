<script setup>
import { computed } from 'vue';
import { COLS } from '../lib/genkou.js';

// layout() が返す1設問分のブロック
const props = defineProps({ block: { type: Object, required: true } });
const emit = defineEmits(['update:indent']);

const countLabel = computed(() => {
  const b = props.block;
  return b.limit ? `（${b.count}／${b.limit}字）` : `（${b.count}字）`;
});

const cols = Array.from({ length: COLS }, (_, c) => c);

function cellClass(r, c) {
  const cell = props.block.rows[r][c];
  return {
    top: r === 0,
    first: c === 0,
    half: cell?.half,
    excess: props.block.limit && r * COLS + c >= props.block.limit,
  };
}
</script>

<template>
  <div class="block">
    <h2>
      <template v-if="block.title">{{ block.title }} </template>
      <span :class="{ over: block.over }">{{ countLabel }}<template v-if="block.over"> 字数超過</template></span>
    </h2>
    <label class="indent"><input type="checkbox" :checked="block.indent" @change="emit('update:indent', $event.target.checked)" /> 段落の書き出しを1マスあける</label>
    <div class="sheet">
      <template v-for="(row, r) in block.rows" :key="r">
        <div v-for="c in cols" :key="c" class="cell" :class="cellClass(r, c)">
          <template v-if="row[c]">{{ row[c].text }}<span v-if="row[c].hang" class="hang">{{ row[c].hang }}</span></template>
        </div>
        <div class="num">{{ (r + 1) % 5 === 0 ? (r + 1) * COLS : '' }}</div>
      </template>
    </div>
  </div>
</template>

<style scoped>
h2 { font-size: 14px; margin: 0 0 6px; font-weight: 600; }
label.indent { display: block; font-size: 12px; margin: 0 0 6px; }
h2 .over { color: var(--over); }
.sheet { display: grid; grid-template-columns: repeat(20, var(--cell)) 34px; font-family: "Hiragino Mincho ProN", "Yu Mincho", "MS Mincho", serif; }
.cell { width: var(--cell); height: var(--cell); border-right: 1px solid var(--line); border-bottom: 1px solid var(--line);
        display: flex; align-items: center; justify-content: center; font-size: 17px; white-space: nowrap; overflow: visible; }
.cell.top { border-top: 1px solid var(--line); }
.cell.first { border-left: 1px solid var(--line); }
.cell.half { font-size: 15px; letter-spacing: -0.02em; }
.cell.excess { background: #fbe3e0; }
.cell .hang { display: inline-block; width: 0.5em; overflow: visible; }
.num { font-size: 10px; color: #888; display: flex; align-items: flex-end; padding-left: 4px; }
</style>
