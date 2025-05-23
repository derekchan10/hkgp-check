from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os
import re
from pypinyin import lazy_pinyin, Style
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/match', methods=['POST'])
def match_results():
    # 读取Excel文件
    try:
        # 读取账号文件
        account_file = request.files.get('account_file')
        if not account_file:
            return jsonify({
                'status': 'error',
                'message': '请上传账号文件'
            })
        account_df = pd.read_excel(account_file)
        
        # 读取中签结果文件
        result_file = request.files.get('result_file')
        if not result_file:
            return jsonify({
                'status': 'error',
                'message': '请上传中签结果文件'
            })
        result_df = pd.read_excel(result_file)
        
        # 处理账号数据
        account_df.columns = ['name', 'account']
        # 确保账号列为字符串类型
        account_df['account'] = account_df['account'].astype(str).replace(' ', '')
        # 提取账号后三位
        account_df['account_last3'] = account_df['account'].str[-3:].astype(str)
        
        # 处理结果数据
        # 假设结果文件的四列分别是：账号后三位、名字头5个字母，投注数，中签数
        result_df.columns = ['account_last3', 'name_first5', 'buy_count', 'win_count']
        
        # 处理结果数据中的名字，去掉空格
        result_df['name_first5'] = result_df['name_first5'].astype(str).str.replace(' ', '')
        
        # 统计结果文件中不同长度名字的数量，确定最常见的长度
        name_lengths = result_df['name_first5'].str.len().value_counts()
        if not name_lengths.empty:
            # 获取最常见的名字长度
            most_common_length = name_lengths.index[0]
        else:
            # 默认长度为5
            most_common_length = 5
            
        print(f"检测到结果文件中最常见的名字长度为: {most_common_length}")
        
        # 对结果数据中的名字进行大写处理
        result_df['name_first5'] = result_df['name_first5'].str.upper()
        
        # 处理账号文件中的姓名数据，根据检测到的长度来截取
        def process_name(name, target_length):
            name_str = str(name)
            # 检查是否包含英文字母
            if re.search('[a-zA-Z]', name_str):
                # 提取所有英文字母
                letters = re.findall('[a-zA-Z]', name_str)
                return ''.join(letters)[:target_length].upper()
            else:
                # 对中文姓名转换为拼音并大写
                pinyin = lazy_pinyin(name_str, style=Style.NORMAL)
                # 拼接拼音并去掉空格
                full_pinyin = ''.join(pinyin).upper().replace(' ', '')
                return full_pinyin[:target_length]
        
        # 根据检测到的长度来处理账号文件中的姓名
        account_df['name_first5'] = account_df['name'].apply(lambda x: process_name(x, most_common_length))
        
        # 合并数据
        merged_results = []
        total_win_count = 0  # 总计中签数
        for _, account_row in account_df.iterrows():
            # 尝试转换为整数，如果失败则保持字符串格式
            try:
                account_last3 = int(account_row['account_last3'])
            except ValueError:
                account_last3 = account_row['account_last3']
            name_first5 = account_row['name_first5']
            # 在结果文件中查找匹配项
            matches = result_df[(result_df['account_last3'] == account_last3) & 
                              (result_df['name_first5'] == name_first5)]
            
            if not matches.empty:
                for _, match in matches.iterrows():
                    win_count = match['win_count']
                    total_win_count += win_count  # 累加中签数
                    merged_results.append({
                        'account_last3': account_last3,
                        'name_first5': name_first5,
                        'name': account_row['name'],
                        'account': account_row['account'],
                        'buy_count': match['buy_count'],
                        'win_count': win_count
                    })
        
        # 将结果保存到session中，供下载使用
        if merged_results:
            app.config['LAST_RESULTS'] = merged_results
        
        return jsonify({
            'status': 'success',
            'data': merged_results,
            'total_matches': len(merged_results),
            'total_win_count': total_win_count,
            'has_results': bool(merged_results)  # 添加标志位表示是否有结果可供下载
        })
    
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        })

@app.route('/download_csv')
def download_csv():
    try:
        # 获取最后的匹配结果
        results = app.config.get('LAST_RESULTS')
        if not results:
            return jsonify({
                'status': 'error',
                'message': '没有可供下载的数据'
            })
        
        # 创建DataFrame
        df = pd.DataFrame(results)
        
        # 添加总计行
        total_win_count = df['win_count'].sum()
        total_buy_count = df['buy_count'].sum()
        
        # 将DataFrame转换为CSV
        output = BytesIO()
        df.to_csv(output, index=False, encoding='utf-8-sig')  # 使用utf-8-sig以支持Excel中文显示
        output.seek(0)
        
        # 生成下载文件名
        filename = f'中签结果_{pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")}.csv'
        
        return send_file(
            output,
            mimetype='text/csv',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': str(e)
        })

if __name__ == '__main__':
    app.run(debug=True) 