import fs from 'fs';
import path from 'path';
import ExcelJS from 'ExcelJS';
// import klaw from 'klaw';
// const path = require('path');
// const ExcelJS = require('exceljs');
// const klaw = require('klaw');
// import * as klaw from 'klaw'
import fg from 'fast-glob'


function formatBytes(bytes: number) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function exportToExcel2(dirPath:string, outputPath:string) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Files');

    // 添加表头
    worksheet.columns = [
        { header: '文件名', key: 'name' },
        { header: '路径', key: 'path' },
        { header: '格式', key: 'format' },
        { header: '大小', key: 'size' },
    ];
    const entries = await fg('**', { cwd: dirPath, stats: true });
    for (const entry of entries) {
        if (entry.stats && entry.stats.size) {
            worksheet.addRow({
                path: entry.path,
                size: formatBytes(entry.stats.size),
            });
        }
    }
    // for await (const file of klaw(dirPath)) {
    //     const filePath = file.path;
    //     const stats = fs.statSync(filePath);

    //     if (stats.isFile()) {
    //         const format = path.extname(filePath).slice(1); // 获取文件格式
    //         const size = formatBytes(stats.size); // 获取文件大小

    //         worksheet.addRow({ name: path.basename(filePath), path: filePath, format: format, size: size });
    //     }
    // }

    // const files = await fs.walk(dirPath);

    // for (const file of files) {
    //     const filePath = file.path;
    //     const stats = await fs.stat(filePath);

    //     if (stats.isFile()) {
    //         const format = path.extname(filePath).slice(1); // 获取文件格式
    //         const size = formatBytes(stats.size); // 获取文件大小

    //         worksheet.addRow({ name: path.basename(filePath), path: filePath, format: format, size: size });
    //     }
    // }

    // fs.readdirSync(dirPath).forEach((file) => {
    //     const filePath = path.join(dirPath, file);
    //     const stats = fs.statSync(filePath);

    //     if (stats.isFile()) {
    //         const format = path.extname(file).slice(1); // 获取文件格式
    //         const size = formatBytes(stats.size); // 获取文件大小

    //         worksheet.addRow({ name: file, path: filePath, format: format, size: size });
    //     }
    // });

    const buffer:any = await workbook.xlsx.writeBuffer();
    fs.writeFileSync(outputPath, buffer);
}

// 使用当前执行路径作为目录路径
// exportToExcel2(process.cwd(), './output.xlsx');

// exportToExcel(process.cwd(), './output.xlsx')

// module.exports = {
//     exportToExcel: exportToExcel2
// }
export default exportToExcel2