const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const XLSX = require('xlsx');
const mammoth = require('mammoth');
const { listDir } = require('../utils');

// 支持的文件扩展名
const WORD_EXTENSIONS = ['.doc', '.docx'];
const EXCEL_EXTENSIONS = ['.xls', '.xlsx', '.xlsm'];
const SUPPORTED_EXTENSIONS = [...WORD_EXTENSIONS, ...EXCEL_EXTENSIONS];

/**
 * 计算文件的MD5哈希值
 * @param {string} filePath 文件路径
 * @returns {string} MD5哈希值
 */
function getFileHash(filePath) {
  try {
    const fileBuffer = fs.readFileSync(filePath);
    const hashSum = crypto.createHash('md5');
    hashSum.update(fileBuffer);
    return hashSum.digest('hex');
  } catch (error) {
    console.error(`Error calculating hash for ${filePath}:`, error.message);
    return null;
  }
}

/**
 * 统计Excel文件的sheet数量
 * @param {string} filePath Excel文件路径
 * @returns {Promise<number>} sheet数量
 */
async function countExcelSheets(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    return workbook.SheetNames.length;
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
    const result = await mammoth.extractRawText({ path: filePath });
    const text = result.value;
    const estimatedPages = Math.max(1, Math.ceil(text.length / 2000));
    return estimatedPages;
  } catch (error) {
    console.error(`Error reading Word file ${filePath}:`, error.message);
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
 * 扫描目录并获取文件信息
 * @param {string} dir 目录路径
 * @param {boolean} recursive 是否递归扫描
 * @returns {Promise<Map>} 文件信息映射表
 */
async function scanDirectory(dir, recursive = false) {
  const fileMap = new Map();
  const items = listDir(dir, recursive).filter(e => !e.isDir);
  
  // 过滤出Office文件
  const officeFiles = items.filter(item => {
    const ext = path.extname(item.name).toLowerCase();
    return SUPPORTED_EXTENSIONS.includes(ext);
  });

  console.log(`Scanning ${officeFiles.length} Office files in ${dir}...`);

  for (let i = 0; i < officeFiles.length; i++) {
    const file = officeFiles[i];
    const progress = `[${i + 1}/${officeFiles.length}]`;
    
    console.log(`${progress} Processing: ${file.name}...`);
    
    const relativePath = path.relative(dir, file.path);
    const hash = getFileHash(file.path);
    const pageCount = await getFilePageCount(file.path);
    
    const ext = path.extname(file.name).toLowerCase();
    const fileType = EXCEL_EXTENSIONS.includes(ext) ? 'Excel' : 'Word';
    const unit = EXCEL_EXTENSIONS.includes(ext) ? 'sheets' : 'pages';
    
    fileMap.set(relativePath, {
      name: file.name,
      path: file.path,
      relativePath: relativePath,
      hash: hash,
      pages: pageCount,
      type: fileType,
      unit: unit,
      size: file.stats.size,
      mtime: file.stats.mtime
    });
  }

  return fileMap;
}

/**
 * 比较两个目录的文件变化
 * @param {string} beforeDir 修改前目录
 * @param {string} afterDir 修改后目录
 * @param {object} opts 选项
 */
module.exports = async function compareDirs(beforeDir, afterDir, opts) {
  const beforePath = path.resolve(beforeDir);
  const afterPath = path.resolve(afterDir);
  
  // 验证目录存在
  if (!fs.existsSync(beforePath) || !fs.statSync(beforePath).isDirectory()) {
    console.error('Before directory not found:', beforePath);
    process.exitCode = 1;
    return;
  }
  
  if (!fs.existsSync(afterPath) || !fs.statSync(afterPath).isDirectory()) {
    console.error('After directory not found:', afterPath);
    process.exitCode = 1;
    return;
  }

  const recursive = !!opts.recursive;

  console.log('='.repeat(80));
  console.log('DIRECTORY COMPARISON');
  console.log('='.repeat(80));
  console.log(`Before: ${beforePath}`);
  console.log(`After:  ${afterPath}`);
  console.log('');

  // 扫描两个目录
  console.log('📁 Scanning BEFORE directory...');
  const beforeFiles = await scanDirectory(beforePath, recursive);
  
  console.log('');
  console.log('📁 Scanning AFTER directory...');
  const afterFiles = await scanDirectory(afterPath, recursive);

  console.log('');
  console.log('🔍 Analyzing changes...');

  const newFiles = [];
  const modifiedFiles = [];
  const deletedFiles = [];
  const unchangedFiles = [];
  let newPagesCount = 0;
  let modifiedPagesCount = 0;
  let deletedPagesCount = 0;

  // Calculate total pages for before and after directories
  let totalBeforePages = 0;
  let totalAfterPages = 0;

  for (const [, info] of beforeFiles) {
    totalBeforePages += info.pages;
  }

  for (const [, info] of afterFiles) {
    totalAfterPages += info.pages;
  }

  // 检查新增和修改的文件
  for (const [relativePath, afterFile] of afterFiles) {
    if (!beforeFiles.has(relativePath)) {
      // 新增文件
      newFiles.push(afterFile);
      newPagesCount += afterFile.pages;
    } else {
      const beforeFile = beforeFiles.get(relativePath);
      if (beforeFile.hash !== afterFile.hash) {
        // 修改的文件
        const pageChange = afterFile.pages - beforeFile.pages;
        modifiedFiles.push({
          ...afterFile,
          beforePages: beforeFile.pages,
          afterPages: afterFile.pages,
          pagesDiff: pageChange
        });
        modifiedPagesCount += pageChange;
      } else {
        // 未变化文件
        unchangedFiles.push(afterFile);
      }
    }
  }

  // 检查删除的文件
  for (const [relativePath, beforeFile] of beforeFiles) {
    if (!afterFiles.has(relativePath)) {
      deletedFiles.push(beforeFile);
      deletedPagesCount += beforeFile.pages;
    }
  }

  // 显示结果
  console.log('='.repeat(80));
  console.log('COMPARISON RESULTS');
  console.log('='.repeat(80));

  if (newFiles.length > 0) {
    console.log(`\n✅ NEW FILES (${newFiles.length}):`);
    console.log('-'.repeat(60));
    newFiles.forEach(file => {
      console.log(`+ ${file.relativePath.padEnd(40)} | ${file.pages} ${file.unit}`);
    });
  }

  if (modifiedFiles.length > 0) {
    console.log(`\n📝 MODIFIED FILES (${modifiedFiles.length}):`);
    console.log('-'.repeat(60));
    modifiedFiles.forEach(file => {
      const diffStr = file.pagesDiff > 0 ? `+${file.pagesDiff}` : `${file.pagesDiff}`;
      console.log(`~ ${file.relativePath.padEnd(40)} | ${file.beforePages} → ${file.afterPages} ${file.unit} (${diffStr})`);
    });
  }

  if (deletedFiles.length > 0) {
    console.log(`\n❌ DELETED FILES (${deletedFiles.length}):`);
    console.log('-'.repeat(60));
    deletedFiles.forEach(file => {
      console.log(`- ${file.relativePath.padEnd(40)} | ${file.pages} ${file.unit}`);
    });
  }

  // 统计信息
  console.log('\n' + '='.repeat(80));
  console.log('SUMMARY');
  console.log('='.repeat(80));
  
  const totalNewPages = newFiles.reduce((sum, f) => sum + f.pages, 0);
  const totalPagesDiff = modifiedFiles.reduce((sum, f) => sum + f.pagesDiff, 0);
  const totalDeletedPages = deletedFiles.reduce((sum, f) => sum + f.pages, 0);

  console.log(`Total files in BEFORE: ${beforeFiles.size} (${totalBeforePages} pages)`);
  console.log(`Total files in AFTER:  ${afterFiles.size} (${totalAfterPages} pages)`);
  console.log(`New files:              ${newFiles.length} (${totalNewPages} pages)`);
  console.log(`Modified files:         ${modifiedFiles.length} (${totalPagesDiff > 0 ? '+' : ''}${totalPagesDiff} pages)`);
  console.log(`Deleted files:          ${deletedFiles.length} (${totalDeletedPages} pages)`);
  console.log(`Unchanged files:        ${unchangedFiles.length}`);
  
  const netPageChange = totalNewPages + totalPagesDiff - totalDeletedPages;
  console.log(`Net page change:        ${netPageChange > 0 ? '+' : ''}${netPageChange} pages`);
};