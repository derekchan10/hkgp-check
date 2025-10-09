document.addEventListener('DOMContentLoaded', function() {
    const matchForm = document.getElementById('matchForm');
    const resultSection = document.getElementById('resultSection');
    const resultTable = document.getElementById('resultTable');
    const totalMatchesElement = document.getElementById('totalMatches');
    const totalUnmatchedElement = document.getElementById('totalUnmatched');
    const summarySection = document.getElementById('summarySection');
    const totalWinCountElement = document.getElementById('totalWinCount');
    const downloadBtn = document.getElementById('downloadBtn');

    // 存储当前匹配结果
    let currentResults = null;

    // 将数据转换为 CSV 格式
    function convertToCSV(data) {
        if (!data || data.length === 0) return '';

        // CSV 表头
        const headers = ['姓名', '账号', '投注数', '中签数', '状态'];
        const headerLine = headers.join(',');

        // CSV 数据行
        const dataLines = data.map(item => {
            return [
                item.name,
                item.account,
                item.buy_count,
                item.win_count,
                item.status === 'matched' ? '已匹配' : '未匹配'
            ].join(',');
        });

        // 添加 BOM 以支持 Excel 中文显示
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
    
    matchForm.addEventListener('submit', function(e) {
        e.preventDefault();
        
        // 检查文件是否已选择
        const accountFile = document.getElementById('accountFile').files[0];
        const resultFile = document.getElementById('resultFile').files[0];
        
        if (!accountFile) {
            alert('请选择账号文件');
            return;
        }
        
        if (!resultFile) {
            alert('请选择中签结果文件');
            return;
        }
        
        // 显示加载状态
        const submitBtn = matchForm.querySelector('button[type="submit"]');
        const originalBtnText = submitBtn.innerHTML;
        submitBtn.innerHTML = '<span class="loading"></span>处理中...';
        submitBtn.disabled = true;
        
        // 准备表单数据
        const formData = new FormData(matchForm);
        
        // 发送请求
        fetch('/match', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            // 恢复按钮状态
            submitBtn.innerHTML = originalBtnText;
            submitBtn.disabled = false;
            
            // 处理响应
            if (data.status === 'success') {
                // 保存当前匹配结果
                currentResults = data.data;

                // 清空之前的结果
                resultTable.innerHTML = '';

                // 显示结果区域
                resultSection.style.display = 'block';

                // 设置匹配总数
                totalMatchesElement.textContent = `匹配: ${data.total_matches}`;
                totalUnmatchedElement.textContent = `未匹配: ${data.total_unmatched}`;

                // 如果没有数据
                if (data.total_matches === 0 && data.total_unmatched === 0) {
                    const noResultRow = document.createElement('tr');
                    noResultRow.innerHTML = '<td colspan="5" class="text-center">没有数据</td>';
                    resultTable.appendChild(noResultRow);
                    summarySection.style.display = 'none';
                    downloadBtn.style.display = 'none';
                    currentResults = null;
                    return;
                }
                
                // 添加结果到表格
                data.data.forEach(item => {
                    const row = document.createElement('tr');
                    
                    // 根据状态设置行的样式
                    if (item.status === 'matched') {
                        row.classList.add('result-highlight');
                    } else {
                        row.classList.add('unmatched-row');
                    }
                    
                    // 设置状态显示
                    let statusHtml = '';
                    if (item.status === 'matched') {
                        statusHtml = '<span class="badge bg-success">已匹配</span>';
                    } else {
                        statusHtml = '<span class="badge bg-warning">未匹配</span>';
                    }
                    
                    row.innerHTML = `
                        <td>${item.name}</td>
                        <td>${item.account}</td>
                        <td>${item.buy_count || '-'}</td>
                        <td>${item.win_count > 0 ? '<span class="text-success fw-bold">' + item.win_count + '</span>' : (item.status === 'matched' ? '0' : '-')}</td>
                        <td>${statusHtml}</td>
                    `;
                    
                    resultTable.appendChild(row);
                });
                
                // 显示总中签数
                if (data.total_win_count > 0) {
                    totalWinCountElement.textContent = data.total_win_count;
                    summarySection.style.display = 'block';
                } else {
                    summarySection.style.display = 'none';
                }
                
                // 显示下载按钮
                downloadBtn.style.display = data.has_results ? 'inline-block' : 'none';
                
                // 平滑滚动到结果区域
                resultSection.scrollIntoView({ behavior: 'smooth' });
            } else {
                // 显示错误信息
                alert('错误: ' + data.message);
            }
        })
        .catch(error => {
            // 恢复按钮状态
            submitBtn.innerHTML = originalBtnText;
            submitBtn.disabled = false;
            
            // 显示错误信息
            alert('请求出错: ' + error);
        });
    });
}); 