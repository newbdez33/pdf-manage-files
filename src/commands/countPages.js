const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const mammoth = require('mammoth');
const { listDir } = require('../utils');

// 支持的文件扩展名
const WORD_EXTENSIONS = ['.doc', '.docx'];
const EXCEL_EXTENSIONS = ['.xls', '.xlsx', '.xlsm'];
const SUPPORTED_EXTENSIONS = [...WORD_EXTENSIONS, ...EXCEL_EXTENSIONS];

/**
 * 统计Excel文件的sheet数量（页数）
 * @param {string} filePath Excel文件路径
 * @returns {Promise<number>} sheet数量
 */
async function countExcelSheets(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheetCount = workbook.SheetNames.length;
    return sheetCount;
  } catch (error) {
    console.error(`Error reading Excel file ${filePath}:`, error.message);
    return 0;
  }
}

/**
 * 统计Word文件的页数
 * @param {string} filePath Word文件路径
 * @returns {Promise<number>} 页数
 */
async function countWordPages(filePath) {
  try {
    // 对于Word文件，我们使用mammoth来提取文本内容
    // 然后根据内容长度估算页数（这是一个近似值）
    const result = await mammoth.extractRawText({ path: filePath });
    const text = result.value;
    
    // 简单的页数估算：假设每页约2000个字符（包括空格）
    // 这是一个粗略的估算，实际页数可能因格式、字体大小等而异
    const estimatedPages = Math.max(1, Math.ceil(text.length / 2000));
    
    return estimatedPages;
  } catch (error) {
    console.error(`Error reading Word file ${filePath}:`, error.message);
    // 如果无法读取，返回1作为默认值
    return 1;
  }
}

/**
 * 获取文件的页数
 * @param {string} filePath 文件路径
 * @returns {Promise<number>} 页数
 */
async function getFilePageCount(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  
  if (EXCEL_EXTENSIONS.includes(ext)) {
    return await countExcelSheets(filePath);
  } else if (WORD_EXTENSIONS.includes(ext)) {
    return await countWordPages(filePath);
  }
  
  return 0;
}

/**
 * 主函数：统计目录中Word/Excel文档的页数
 * @param {string} dir 目录路径
 * @param {object} opts 选项
 */
module.exports = async function countPages(dir, opts) {
  const targetDir = path.resolve(dir);
  
  if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
    console.error('Directory not found:', targetDir);
    process.exitCode = 1;
    return;
  }

  const recursive = !!opts.recursive;
  const items = listDir(targetDir, recursive).filter(e => !e.isDir);
  
  // 过滤出Word和Excel文件
  const officeFiles = items.filter(item => {
    const ext = path.extname(item.name).toLowerCase();
    return SUPPORTED_EXTENSIONS.includes(ext);
  });

  if (officeFiles.length === 0) {
    console.log('No Word or Excel files found in the directory.');
    return;
  }

  console.log(`Found ${officeFiles.length} Office document(s) in ${targetDir}`);
  console.log('');

  let totalPages = 0;
  const results = [];

  // 统计每个文件的页数
  for (let i = 0; i < officeFiles.length; i++) {
    const file = officeFiles[i];
    
    // 显示进度
    const progress = `[${i + 1}/${officeFiles.length}]`;
    console.log(`${progress} Processing: ${file.name}...`);
    
    const pageCount = await getFilePageCount(file.path);
    totalPages += pageCount;
    
    const ext = path.extname(file.name).toLowerCase();
    const fileType = EXCEL_EXTENSIONS.includes(ext) ? 'Excel' : 'Word';
    const unit = EXCEL_EXTENSIONS.includes(ext) ? 'sheets' : 'pages';
    
    results.push({
      path: file.path,
      name: file.name,
      type: fileType,
      pages: pageCount,
      unit: unit
    });
    
    // 显示当前文件的结果
    console.log(`${progress} ${file.name} - ${pageCount} ${unit}`);
  }

  console.log('');
  console.log('='.repeat(80));
  console.log('SUMMARY - Page count results:');
  console.log('='.repeat(80));
  
  results.forEach(result => {
    const relativePath = path.relative(targetDir, result.path);
    console.log(`${result.name.padEnd(40)} | ${result.type.padEnd(5)} | ${result.pages} ${result.unit}`);
  });
  
  console.log('='.repeat(80));
  console.log(`Total files: ${officeFiles.length}`);
  console.log(`Total pages: ${totalPages}`);
  
  // 按文件类型分组统计
  const wordFiles = results.filter(r => r.type === 'Word');
  const excelFiles = results.filter(r => r.type === 'Excel');
  
  if (wordFiles.length > 0) {
    const wordPages = wordFiles.reduce((sum, f) => sum + f.pages, 0);
    console.log(`Word files: ${wordFiles.length} files, ${wordPages} pages`);
  }
  
  if (excelFiles.length > 0) {
    const excelSheets = excelFiles.reduce((sum, f) => sum + f.pages, 0);
    console.log(`Excel files: ${excelFiles.length} files, ${excelSheets} sheets`);
  }
};