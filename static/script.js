document.addEventListener('DOMContentLoaded', function() {
    const matchForm = document.getElementById('matchForm');
    const resultSection = document.getElementById('resultSection');
    const resultTable = document.getElementById('resultTable');
    const totalMatchesElement = document.getElementById('totalMatches');
    const summarySection = document.getElementById('summarySection');
    const totalWinCountElement = document.getElementById('totalWinCount');
    
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
                // 清空之前的结果
                resultTable.innerHTML = '';
                
                // 显示结果区域
                resultSection.style.display = 'block';
                
                // 设置匹配总数
                totalMatchesElement.textContent = data.total_matches;
                
                // 如果没有匹配项
                if (data.total_matches === 0) {
                    const noResultRow = document.createElement('tr');
                    noResultRow.innerHTML = '<td colspan="4" class="text-center">未找到匹配项</td>';
                    resultTable.appendChild(noResultRow);
                    summarySection.style.display = 'none';
                    return;
                }
                
                // 添加结果到表格
                data.data.forEach(item => {
                    const row = document.createElement('tr');
                    row.classList.add('result-highlight');
                    
                    row.innerHTML = `
                        <td>${item.name}</td>
                        <td>${item.account}</td>
                        <td>${item.buy_count}</td>
                        <td>${item.win_count > 0 ? '<span class="text-success fw-bold">' + item.win_count + '</span>' : '0'}</td>
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