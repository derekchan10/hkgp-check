document.addEventListener('DOMContentLoaded', function() {
    const matchForm = document.getElementById('matchForm');
    const resultSection = document.getElementById('resultSection');
    const resultTable = document.getElementById('resultTable');
    const summarySection = document.getElementById('summarySection');
    const totalWinCountElement = document.getElementById('totalWinCount');
    const downloadBtn = document.getElementById('downloadBtn');

    // 存储当前匹配结果
    let currentResults = null;

    // 将数据转换为 CSV 格式
    function convertToCSV(data) {
        if (!data || data.length === 0) return '';

        const headers = ['姓名', '账号', '投注数', '中签数', '状态'];
        const headerLine = headers.join(',');

        const dataLines = data.map(item => {
            return [
                item.name,
                item.account,
                item.buy_count,
                item.win_count,
                item.status === 'matched' ? '已匹配' : '未匹配'
            ].join(',');
        });

        return '\ufeff' + [headerLine, ...dataLines].join('\n');
    }

    // 下载 CSV 文件
    function downloadCSV(csvContent, filename) {
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);

        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        link.style.visibility = 'hidden';

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        URL.revokeObjectURL(url);
    }

    // 下载按钮点击事件
    downloadBtn.addEventListener('click', function() {
        if (!currentResults || currentResults.length === 0) {
            alert('没有可供下载的数据');
            return;
        }

        const csv = convertToCSV(currentResults);
        const date = new Date().toISOString().split('T')[0].replace(/-/g, '');
        const filename = `中签结果_${date}.csv`;

        downloadCSV(csv, filename);
    });

    // ── 纯前端匹配逻辑 ──────────────────────────────────────────────────

    // 用 FileReader + SheetJS 读取 Excel 文件，返回二维数组（含表头行）
    function readExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                    resolve(rows);
                } catch (err) {
                    reject(new Error('读取文件失败：' + err.message));
                }
            };
            reader.onerror = () => reject(new Error('文件读取出错'));
            reader.readAsArrayBuffer(file);
        });
    }

    // 特殊拼音映射表（与 app.py 保持一致）
    const SPECIAL_PINYIN = {
        '长': 'chang',
        '翟': 'zhai',
        '绿': 'lyu',
        '曾': 'zeng',
        '靓': 'liang',
        '吕': 'lyu'
    };

    // 姓名处理：中文转拼音 / 英文提取字母，与 app.py process_name() 逻辑一致
    function processName(name, targetLength) {
        const nameStr = String(name).trim();
        if (/[a-zA-Z]/.test(nameStr)) {
            // 英文名：提取所有字母，大写，截取
            return nameStr.replace(/[^a-zA-Z]/g, '').toUpperCase().slice(0, targetLength);
        } else {
            // 中文名：逐字转拼音，优先使用特殊映射
            let result = '';
            for (const char of nameStr) {
                if (SPECIAL_PINYIN[char] !== undefined) {
                    result += SPECIAL_PINYIN[char];
                } else {
                    const py = pinyinPro.pinyin(char, { toneType: 'none' });
                    result += py.replace(/\s/g, '');
                }
            }
            return result.toUpperCase().slice(0, targetLength);
        }
    }

    // 检测结果文件中最常见的名字前缀长度（与 app.py 逻辑一致）
    function detectMostCommonLength(nameList) {
        const lengthCount = {};
        for (const name of nameList) {
            const len = name.length;
            lengthCount[len] = (lengthCount[len] || 0) + 1;
        }
        let mostCommonLen = 5;
        let maxCount = 0;
        for (const [len, count] of Object.entries(lengthCount)) {
            if (count > maxCount) {
                maxCount = count;
                mostCommonLen = parseInt(len);
            }
        }
        return mostCommonLen;
    }

    // 核心匹配逻辑（移植自 app.py match_results()）
    function matchData(accountRows, resultRows) {
        // 跳过表头行，解析账号数据
        const accounts = accountRows.slice(1)
            .map(row => ({
                name: String(row[0] || '').trim(),
                account: String(row[1] || '').trim()
            }))
            .filter(a => a.name && a.account);

        if (accounts.length === 0) {
            throw new Error('账号文件中没有有效数据，请检查文件格式（第一行为表头，第一列姓名，第二列账号）');
        }

        // 跳过表头行，解析中签结果数据
        const results = resultRows.slice(1)
            .map(row => ({
                account_last3: row[0],
                name_first5: String(row[1] || '').replace(/\s/g, '').toUpperCase(),
                buy_count: row[2] || 0,
                win_count: row[3] || 0
            }))
            .filter(r => r.account_last3 !== '' && r.name_first5 !== '');

        if (results.length === 0) {
            throw new Error('中签结果文件中没有有效数据，请检查文件格式');
        }

        // 将 account_last3 转为整数
        results.forEach(r => {
            r.account_last3 = parseInt(r.account_last3);
        });

        // 检测最常见名字长度
        const mostCommonLength = detectMostCommonLength(results.map(r => r.name_first5));

        // 构建快速查找 Map：key = `${account_last3}_${name_first5}`
        const resultMap = new Map();
        for (const r of results) {
            const key = `${r.account_last3}_${r.name_first5}`;
            if (!resultMap.has(key)) {
                resultMap.set(key, []);
            }
            resultMap.get(key).push(r);
        }

        // 遍历所有账号，执行匹配
        const mergedResults = [];
        let totalWinCount = 0;

        for (const acc of accounts) {
            const accountLast3 = parseInt(acc.account.slice(-3));
            const nameFirst5 = processName(acc.name, mostCommonLength);
            const key = `${accountLast3}_${nameFirst5}`;
            const matches = resultMap.get(key);

            if (matches && matches.length > 0) {
                for (const match of matches) {
                    const winCount = Number(match.win_count);
                    totalWinCount += winCount;
                    mergedResults.push({
                        account_last3: accountLast3,
                        name_first5: nameFirst5,
                        name: acc.name,
                        account: acc.account,
                        buy_count: match.buy_count,
                        win_count: winCount,
                        status: 'matched'
                    });
                }
            } else {
                mergedResults.push({
                    account_last3: accountLast3,
                    name_first5: nameFirst5,
                    name: acc.name,
                    account: acc.account,
                    buy_count: 0,
                    win_count: 0,
                    status: 'unmatched'
                });
            }
        }

        return {
            status: 'success',
            data: mergedResults,
            total_matches: mergedResults.filter(r => r.status === 'matched').length,
            total_unmatched: mergedResults.filter(r => r.status === 'unmatched').length,
            total_win_count: totalWinCount,
            has_results: mergedResults.length > 0
        };
    }

    // 渲染结果（与原逻辑完全相同）
    function renderResults(data) {
        currentResults = data.data;

        resultTable.innerHTML = '';
        resultSection.style.display = 'block';

        if (data.total_matches === 0 && data.total_unmatched === 0) {
            const noResultRow = document.createElement('tr');
            noResultRow.innerHTML = '<td colspan="5" class="text-center">没有数据</td>';
            resultTable.appendChild(noResultRow);
            summarySection.style.display = 'none';
            downloadBtn.style.display = 'none';
            currentResults = null;
            return;
        }

        data.data.forEach(item => {
            const row = document.createElement('tr');

            if (item.status === 'matched') {
                row.classList.add('result-highlight');
            } else {
                row.classList.add('unmatched-row');
            }

            const statusHtml = item.status === 'matched'
                ? '<span class="badge bg-success">已匹配</span>'
                : '<span class="badge bg-warning">未匹配</span>';

            row.innerHTML = `
                <td>${item.name}</td>
                <td>${item.account}</td>
                <td>${item.buy_count || '-'}</td>
                <td>${item.win_count > 0 ? '<span class="text-success fw-bold">' + item.win_count + '</span>' : (item.status === 'matched' ? '0' : '-')}</td>
                <td>${statusHtml}</td>
            `;

            resultTable.appendChild(row);
        });

        document.getElementById('totalMatchesSummary').textContent = data.total_matches;
        document.getElementById('totalUnmatchedSummary').textContent = data.total_unmatched;
        totalWinCountElement.textContent = data.total_win_count;
        const wonAccounts = data.data.filter(i => i.status === 'matched' && i.win_count > 0).length;
        const rate = data.total_matches > 0 ? (wonAccounts / data.total_matches * 100).toFixed(1) + '%' : '-';
        document.getElementById('totalWinRate').textContent = rate;
        summarySection.style.display = data.has_results ? 'block' : 'none';

        downloadBtn.style.display = data.has_results ? 'inline-block' : 'none';

        resultSection.scrollIntoView({ behavior: 'smooth' });
    }

    // ── 表单提交：本地处理，无需上传 ────────────────────────────────────

    matchForm.addEventListener('submit', async function(e) {
        e.preventDefault();

        const accountFile = document.getElementById('accountFile').files[0];
        const resultFile = document.getElementById('resultFile').files[0];

        if (!accountFile) { alert('请选择账号文件'); return; }
        if (!resultFile) { alert('请选择中签结果文件'); return; }

        const submitBtn = matchForm.querySelector('button[type="submit"]');
        const originalBtnText = submitBtn.innerHTML;
        submitBtn.innerHTML = '<span class="loading"></span>处理中...';
        submitBtn.disabled = true;

        try {
            // 并行读取两个文件
            const [accountRows, resultRows] = await Promise.all([
                readExcelFile(accountFile),
                readExcelFile(resultFile)
            ]);

            const data = matchData(accountRows, resultRows);
            renderResults(data);
        } catch (error) {
            alert('处理出错: ' + error.message);
        } finally {
            submitBtn.innerHTML = originalBtnText;
            submitBtn.disabled = false;
        }
    });
});