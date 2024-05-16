<script setup lang="ts">
import { ref } from 'vue'
// import path from 'path'
// import { dialog, app, shell } from '@electron/remote'
// const { dialog, app, shell } = require('@electron/remote')
// import exportToExcel from '../utils/exportExcel'
// const { exportToExcel } = require('../utils/exportExcel')

defineProps<{ msg: string }>()

const exportExcel = async () => {
  await window.ipcRenderer.invoke('export-excel')
  // const _pathes = await dialog.showOpenDialog({
  //     properties: ['openDirectory'],
  // })
  // console.error(_pathes)
  // const xPath = path.join(app.getPath('pictures'), '/output.xlsx')
  // exportToExcel(_pathes[0], xPath)
  // exportToExcel(process.cwd(), './output.xlsx')
  // shell.openItem(app.getPath('pictures'))
}
</script>

<template>
    <h1>{{ msg }}</h1>

    <div class="card">
        <button type="button" @click="exportExcel">选择文件夹导出</button>
    </div>
</template>

<style scoped>
.read-the-docs {
    color: #888;
}
</style>
