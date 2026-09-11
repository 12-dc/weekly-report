# -*- coding: utf-8 -*-
"""
周报助手 - 后端（仅做Excel填入，OCR在前端浏览器完成）
"""
import os
import sys
import json
import uuid
import threading
import time
from flask import Flask, request, send_file, jsonify, Response

app = Flask(__name__)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from excel_filler import ExcelFiller

TEMP_DIR = os.environ.get('TEMP_DIR', os.path.join(SCRIPT_DIR, 'temp'))
os.makedirs(TEMP_DIR, exist_ok=True)


@app.route('/')
def index():
    return Response(HTML_PAGE, mimetype='text/html; charset=utf-8')


@app.route('/status')
def status():
    return jsonify({'status': 'ok'})


@app.route('/fill', methods=['POST'])
def fill():
    try:
        target_name = request.form.get('target_name', '陈子怡').strip() or '陈子怡'

        template_file = request.files.get('template')
        if not template_file:
            return jsonify({'error': '请上传周报模板'}), 400

        parsed_data_str = request.form.get('parsed_data', '[]')
        try:
            parsed_results = json.loads(parsed_data_str)
        except Exception:
            return jsonify({'error': '解析数据格式错误'}), 400

        if not parsed_results:
            return jsonify({'error': '没有识别到工作内容'}), 400

        # 保存模板
        template_path = os.path.join(TEMP_DIR, f'template_{uuid.uuid4().hex}.xlsx')
        template_file.save(template_path)

        # 加载配置
        config_path = os.path.join(SCRIPT_DIR, 'config.json')
        team_info = ''
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            team_info = cfg.get('team_info', '')

        # 生成Excel
        output_filename = f'report_{uuid.uuid4().hex[:8]}.xlsx'
        output_path = os.path.join(TEMP_DIR, output_filename)

        filler = ExcelFiller(
            template_path=template_path,
            output_path=output_path,
            team_info=team_info
        )
        filler.fill(parsed_results)

        # 清理模板
        try:
            os.remove(template_path)
        except Exception:
            pass

        # 10分钟后清理输出文件
        def cleanup():
            time.sleep(600)
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
            except Exception:
                pass
        threading.Thread(target=cleanup, daemon=True).start()

        return jsonify({
            'success': True,
            'download_url': f'/download/{output_filename}',
            'filename': output_filename
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': f'生成失败：{str(e)}'}), 500


@app.route('/download/<filename>')
def download(filename):
    filepath = os.path.join(TEMP_DIR, filename)
    if os.path.exists(filepath):
        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return jsonify({'error': '文件不存在'}), 404


HTML_PAGE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>周报助手</title>
<script src="https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js"></script>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
}
.container {
    background: #fff;
    border-radius: 16px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    width: 100%;
    max-width: 680px;
    padding: 40px;
}
h1 { font-size: 28px; color: #333; margin-bottom: 8px; text-align: center; }
.subtitle { text-align: center; color: #888; margin-bottom: 32px; font-size: 14px; }
.form-group { margin-bottom: 24px; }
label { display: block; font-weight: 600; color: #555; margin-bottom: 8px; font-size: 14px; }
.file-input-wrapper {
    position: relative; border: 2px dashed #ddd; border-radius: 10px; padding: 24px;
    text-align: center; cursor: pointer; transition: all 0.3s; background: #fafafa;
}
.file-input-wrapper:hover { border-color: #667eea; background: #f0f0ff; }
.file-input-wrapper.dragover { border-color: #667eea; background: #e8e8ff; }
.file-input-wrapper input[type=file] {
    position: absolute; top: 0; left: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer;
}
.file-icon { font-size: 36px; margin-bottom: 8px; }
.file-text { color: #666; font-size: 14px; }
.file-list { margin-top: 12px; font-size: 13px; color: #667eea; }
.file-list span { display: inline-block; background: #f0f0ff; padding: 2px 10px; border-radius: 12px; margin: 2px; }
input[type=text] {
    width: 100%; padding: 12px 16px; border: 1px solid #ddd; border-radius: 8px; font-size: 15px;
}
input[type=text]:focus { outline: none; border-color: #667eea; }
.btn {
    width: 100%; padding: 14px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: #fff; border: none; border-radius: 10px; font-size: 16px; font-weight: 600; cursor: pointer;
    transition: transform 0.2s, opacity 0.3s;
}
.btn:hover { transform: translateY(-2px); }
.btn:disabled { opacity: 0.6; cursor: not-allowed; transform: none; }
.progress-container { display: none; margin-top: 24px; }
.progress-bar { width: 100%; height: 8px; background: #eee; border-radius: 4px; overflow: hidden; }
.progress-fill {
    height: 100%; background: linear-gradient(90deg, #667eea, #764ba2); width: 0%; transition: width 0.3s;
}
.progress-text { text-align: center; margin-top: 8px; font-size: 13px; color: #888; }
.result { display: none; margin-top: 24px; padding: 20px; background: #f0fff4; border: 1px solid #c6f6d5; border-radius: 10px; }
.result h3 { color: #22543d; margin-bottom: 12px; }
.result-item { font-size: 13px; color: #555; margin-bottom: 6px; padding-left: 16px; }
.download-btn {
    display: inline-block; margin-top: 12px; padding: 10px 24px; background: #38a169;
    color: #fff; text-decoration: none; border-radius: 8px; font-weight: 600;
}
.error { display: none; margin-top: 24px; padding: 16px; background: #fff5f5; border: 1px solid #feb2b2; border-radius: 10px; color: #c53030; font-size: 14px; }
.tip { margin-top: 16px; padding: 12px; background: #fffff0; border: 1px solid #fefcbf; border-radius: 8px; font-size: 12px; color: #975a16; line-height: 1.6; }
</style>
</head>
<body>
<div class="container">
    <h1>周报助手</h1>
    <p class="subtitle">自动识别排期图片中的工作内容，一键填入周报模板</p>

    <div class="form-group">
        <label>周报模板（Excel）</label>
        <div class="file-input-wrapper" id="templateWrapper">
            <input type="file" id="templateInput" accept=".xlsx,.xls">
            <div class="file-icon">📊</div>
            <div class="file-text">点击选择或拖拽Excel模板到此处</div>
            <div class="file-list" id="templateList"></div>
        </div>
    </div>

    <div class="form-group">
        <label>排期图片（可多选）</label>
        <div class="file-input-wrapper" id="imageWrapper">
            <input type="file" id="imageInput" accept="image/*" multiple>
            <div class="file-icon">🖼️</div>
            <div class="file-text">点击选择或拖拽图片到此处（支持多选）</div>
            <div class="file-list" id="imageList"></div>
        </div>
    </div>

    <div class="form-group">
        <label>提取人员姓名</label>
        <input type="text" id="targetName" value="陈子怡" placeholder="请输入要提取的人员姓名">
    </div>

    <button class="btn" id="submitBtn" onclick="submitForm()">开始生成周报</button>

    <div class="progress-container" id="progressContainer">
        <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
        <div class="progress-text" id="progressText">正在处理...</div>
    </div>

    <div class="result" id="resultBox">
        <h3>✅ 生成成功！</h3>
        <div id="resultSummary"></div>
        <a href="#" class="download-btn" id="downloadLink" download>📥 下载周报Excel</a>
    </div>

    <div class="error" id="errorBox"></div>

    <div class="tip">
        💡 OCR在你的浏览器中运行，首次使用会下载中文识别包（约15MB），请耐心等待。图片识别速度取决于你的电脑性能。
    </div>
</div>

<script>
let templateFile = null;
let imageFiles = [];
let ocrWorker = null;

document.getElementById('templateInput').addEventListener('change', function(e) {
    if (e.target.files.length > 0) {
        templateFile = e.target.files[0];
        document.getElementById('templateList').innerHTML = '<span>' + templateFile.name + '</span>';
    }
});

document.getElementById('imageInput').addEventListener('change', function(e) {
    imageFiles = Array.from(e.target.files);
    updateImageList();
});

['templateWrapper', 'imageWrapper'].forEach(function(id) {
    var wrapper = document.getElementById(id);
    var input = wrapper.querySelector('input[type=file]');
    wrapper.addEventListener('dragover', function(e) { e.preventDefault(); wrapper.classList.add('dragover'); });
    wrapper.addEventListener('dragleave', function() { wrapper.classList.remove('dragover'); });
    wrapper.addEventListener('drop', function(e) {
        e.preventDefault(); wrapper.classList.remove('dragover');
        input.files = e.dataTransfer.files;
        input.dispatchEvent(new Event('change'));
    });
});

function updateImageList() {
    var list = document.getElementById('imageList');
    list.innerHTML = imageFiles.length === 0 ? '' : imageFiles.map(function(f) { return '<span>' + f.name + '</span>'; }).join('');
}

function setProgress(percent, text) {
    document.getElementById('progressFill').style.width = percent + '%';
    document.getElementById('progressText').textContent = text;
}

function showError(msg) {
    document.getElementById('errorBox').textContent = '❌ ' + msg;
    document.getElementById('errorBox').style.display = 'block';
    document.getElementById('progressContainer').style.display = 'none';
    document.getElementById('submitBtn').disabled = false;
    document.getElementById('submitBtn').textContent = '开始生成周报';
}

// OCR识别单张图片
async function recognizeImage(file) {
    const { data } = await ocrWorker.recognize(file);
    var blocks = [];
    if (data.words) {
        data.words.forEach(function(w) {
            if (w.text && w.text.trim()) {
                var b = w.bbox;
                blocks.push({
                    text: w.text.trim(),
                    cx: (b.x0 + b.x1) / 2,
                    cy: (b.y0 + b.y1) / 2,
                    x_min: b.x0, x_max: b.x1,
                    y_min: b.y0, y_max: b.y1,
                    height: b.y1 - b.y0
                });
            }
        });
    }
    return blocks;
}

// 解析表格，提取目标人员工作内容
function parseBlocks(blocks, targetName) {
    if (!blocks || blocks.length === 0) {
        return { found: false, reason: 'OCR未识别到文本' };
    }

    // 检测列边界
    var headerKeywords = ['日期', '人员', '工作内容', '完成时间'];
    var headerBlocks = [];
    blocks.forEach(function(b) {
        for (var i = 0; i < headerKeywords.length; i++) {
            if (b.text.indexOf(headerKeywords[i]) !== -1) {
                headerBlocks.push(b);
                break;
            }
        }
    });

    var columns = {};
    if (headerBlocks.length > 0) {
        headerBlocks.sort(function(a, b) { return a.cx - b.cx; });
        var personHb = null, deadlineHb = null;
        headerBlocks.forEach(function(hb) {
            if (hb.text.indexOf('人员') !== -1) personHb = hb;
            if (hb.text.indexOf('完成时间') !== -1) deadlineHb = hb;
        });

        var personLeft = 0, personRight = 0;
        if (personHb) {
            var px = personHb.cx;
            var nameBlocks = blocks.filter(function(b) {
                return Math.abs(b.cx - px) < 70 && b.text !== '人员' && b.text.length <= 5;
            });
            if (nameBlocks.length > 0) {
                personLeft = Math.min.apply(null, nameBlocks.map(function(b) { return b.x_min; })) - 20;
                personRight = Math.max.apply(null, nameBlocks.map(function(b) { return b.x_max; })) + 25;
            } else {
                personLeft = px - 50; personRight = px + 80;
            }
        }

        var deadlineLeft = 99999;
        if (deadlineHb) {
            var dx = deadlineHb.cx;
            var dlBlocks = blocks.filter(function(b) {
                return Math.abs(b.cx - dx) < 100 && b.text !== '完成时间节点' && b.text !== '完成时间';
            });
            if (dlBlocks.length > 0) {
                deadlineLeft = Math.min.apply(null, dlBlocks.map(function(b) { return b.x_min; })) - 15;
            } else {
                deadlineLeft = dx - 60;
            }
        }

        columns = {
            date: [0, personLeft],
            person: [personLeft, personRight],
            content: [personRight, deadlineLeft],
            deadline: [deadlineLeft, 99999]
        };
    } else {
        // fallback：简单分4列
        var xs = blocks.map(function(b) { return b.cx; }).sort(function(a, b) { return a - b; });
        var n = xs.length;
        var q1 = xs[Math.floor(n/4)], q2 = xs[Math.floor(n/2)], q3 = xs[Math.floor(3*n/4)];
        columns = {
            date: [0, (q1+q2)/2],
            person: [(q1+q2)/2, (q2+q3)/2],
            content: [(q2+q3)/2, (q3+xs[xs.length-1])/2],
            deadline: [(q3+xs[xs.length-1])/2, 99999]
        };
    }

    // 按列分类
    var colBlocks = { date: [], person: [], content: [], deadline: [] };
    blocks.forEach(function(b) {
        for (var name in columns) {
            var range = columns[name];
            if (b.cx >= range[0] && b.cx < range[1]) {
                colBlocks[name].push(b);
                break;
            }
        }
    });

    // 找目标人员
    var personBlocks = colBlocks.person.filter(function(b) { return b.text !== '人员'; })
        .sort(function(a, b) { return a.cy - b.cy; });

    if (personBlocks.length === 0) {
        return { found: false, reason: '未识别到人员列' };
    }

    var targetBlock = null;
    for (var i = 0; i < personBlocks.length; i++) {
        if (personBlocks[i].text.indexOf(targetName) !== -1) {
            targetBlock = personBlocks[i];
            break;
        }
    }
    if (!targetBlock) {
        return { found: false, reason: '未找到' + targetName };
    }

    var targetY = targetBlock.cy;
    var personYs = personBlocks.map(function(b) { return b.cy; });
    var idx = personYs.indexOf(targetY);
    if (idx === -1) idx = 0;

    var avgGap = 40;
    if (personYs.length >= 2) {
        var gaps = [];
        for (var i = 0; i < personYs.length - 1; i++) gaps.push(personYs[i+1] - personYs[i]);
        avgGap = gaps.reduce(function(a, b) { return a + b; }, 0) / gaps.length;
    }

    var yMin = idx > 0 ? (personYs[idx-1] + targetY) / 2 : 0;
    var yMax = idx < personYs.length - 1 ? (targetY + personYs[idx+1]) / 2 : targetY + avgGap * 1.5;

    // 收集工作内容
    var contentBlocks = colBlocks.content.filter(function(b) {
        return b.cy >= yMin && b.cy <= yMax;
    }).sort(function(a, b) { return a.cy - b.cy || a.cx - b.cx; });

    // 过滤噪声
    contentBlocks = contentBlocks.filter(function(b) {
        if (!b.text || b.text.length < 2) return false;
        var valid = 0;
        for (var i = 0; i < b.text.length; i++) {
            var c = b.text.charCodeAt(i);
            if ((c >= 0x4e00 && c <= 0x9fff) || b.text[i].match(/[a-zA-Z0-9]/)) valid++;
        }
        if (valid / b.text.length < 0.4) return false;
        return true;
    });

    // 按y子聚类合并
    var workItems = [];
    if (contentBlocks.length > 0) {
        var heights = contentBlocks.map(function(b) { return b.height; }).filter(function(h) { return h > 0; });
        var avgH = heights.length > 0 ? heights.reduce(function(a,b){return a+b;},0)/heights.length : 18;
        var subThreshold = avgH * 0.7;

        var subRows = [];
        var current = [contentBlocks[0]];
        var currentY = contentBlocks[0].cy;
        for (var i = 1; i < contentBlocks.length; i++) {
            if (Math.abs(contentBlocks[i].cy - currentY) < subThreshold) {
                current.push(contentBlocks[i]);
                currentY = current.reduce(function(s,b){return s+b.cy;},0) / current.length;
            } else {
                subRows.push(current);
                current = [contentBlocks[i]];
                currentY = contentBlocks[i].cy;
            }
        }
        subRows.push(current);

        subRows.forEach(function(sr) {
            sr.sort(function(a, b) { return a.cx - b.cx; });
            var merged = sr.map(function(b) { return b.text; }).join('');
            merged = merged.replace(/^(\d+)(?=[A-Z\u4e00-\u9fff])/, '$1、');
            merged = merged.replace(/^(\d+)>/, '$1、');
            workItems.push(merged);
        });
    }

    // 提取日期
    var dateBlocks = colBlocks.date.filter(function(b) { return b.text !== '日期'; });
    var dateStr = null, weekday = null;
    var fullDate = dateBlocks.map(function(b) { return b.text; }).join(' ');
    var m = fullDate.match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
    if (m) dateStr = m[1] + '/' + m[2] + '/' + m[3];
    else {
        m = fullDate.match(/(\d{1,2})月(\d{1,2})日/);
        if (m) dateStr = m[1] + '月' + m[2] + '日';
    }
    m = fullDate.match(/周([一二三四五六日天])/);
    if (m) weekday = m[1];

    // 完成时间
    var deadlineBlocks = colBlocks.deadline.filter(function(b) {
        return b.cy >= yMin && b.cy <= yMax;
    });
    var deadline = deadlineBlocks.map(function(b) { return b.text; }).join(' ').trim();

    if (!dateStr && deadline) {
        m = deadline.match(/(\d{1,2})月(\d{1,2})日/);
        if (m) dateStr = m[1] + '月' + m[2] + '日';
    }

    return {
        found: true,
        date: dateStr,
        weekday: weekday,
        person: targetName,
        work_items: workItems,
        deadline: deadline
    };
}

// 按日期排序
function sortByDate(results) {
    function dateKey(r) {
        var d = r.date;
        if (!d) {
            var dl = r.deadline || '';
            var m = dl.match(/(\d{1,2})月(\d{1,2})日/);
            if (m) return [2026, parseInt(m[1]), parseInt(m[2])];
            return [9999, 99, 99];
        }
        var m = d.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
        if (m) return [parseInt(m[1]), parseInt(m[2]), parseInt(m[3])];
        m = d.match(/(\d{1,2})月(\d{1,2})日/);
        if (m) return [2026, parseInt(m[1]), parseInt(m[2])];
        return [9999, 99, 99];
    }
    return results.sort(function(a, b) {
        var ka = dateKey(a), kb = dateKey(b);
        for (var i = 0; i < 3; i++) {
            if (ka[i] !== kb[i]) return ka[i] - kb[i];
        }
        return 0;
    });
}

async function submitForm() {
    if (!templateFile) { showError('请先选择周报模板'); return; }
    if (imageFiles.length === 0) { showError('请至少选择一张图片'); return; }

    var btn = document.getElementById('submitBtn');
    btn.disabled = true;
    btn.textContent = '处理中...';
    document.getElementById('resultBox').style.display = 'none';
    document.getElementById('errorBox').style.display = 'none';
    document.getElementById('progressContainer').style.display = 'block';

    try {
        // 初始化OCR
        setProgress(5, '正在加载OCR引擎（首次使用需下载中文包）...');
        if (!ocrWorker) {
            ocrWorker = await Tesseract.createWorker('chi_sim', 1, {
                logger: function(m) {
                    if (m.status === 'recognizing text') {
                        setProgress(10 + m.progress * 20, '正在识别图片...');
                    }
                }
            });
        }

        // 逐张识别
        var parsedResults = [];
        for (var i = 0; i < imageFiles.length; i++) {
            setProgress(30 + (i / imageFiles.length) * 40, '正在识别第 ' + (i+1) + '/' + imageFiles.length + ' 张图片...');
            var blocks = await recognizeImage(imageFiles[i]);
            var targetName = document.getElementById('targetName').value || '陈子怡';
            var result = parseBlocks(blocks, targetName);
            result.image = imageFiles[i].name;
            parsedResults.push(result);
        }

        // 排序
        parsedResults = sortByDate(parsedResults);

        // 发给后端生成Excel
        setProgress(75, '正在生成Excel...');
        var formData = new FormData();
        formData.append('template', templateFile);
        formData.append('target_name', document.getElementById('targetName').value || '陈子怡');
        formData.append('parsed_data', JSON.stringify(parsedResults));

        var response = await fetch('/fill', { method: 'POST', body: formData });
        var data = await response.json();

        if (data.success) {
            setProgress(100, '完成！');
            // 显示结果
            var summaryHtml = '';
            parsedResults.forEach(function(item) {
                if (item.found && item.work_items && item.work_items.length > 0) {
                    summaryHtml += '<div class="result-item">📅 ' + (item.date || '未知日期') + ' (' + (item.weekday || '') + ')：' + item.work_items.length + '项工作</div>';
                } else {
                    summaryHtml += '<div class="result-item">📅 ' + (item.date || '未知日期') + '：未找到（' + (item.reason || '无数据') + '）</div>';
                }
            });
            document.getElementById('resultSummary').innerHTML = summaryHtml;
            document.getElementById('downloadLink').href = data.download_url;
            document.getElementById('downloadLink').download = data.filename;
            document.getElementById('resultBox').style.display = 'block';
            setTimeout(function() { document.getElementById('progressContainer').style.display = 'none'; }, 1000);
        } else {
            showError(data.error || '生成失败');
        }
    } catch (err) {
        showError('处理出错：' + err.message);
    } finally {
        btn.disabled = false;
        btn.textContent = '开始生成周报';
    }
}
</script>
</body>
</html>'''


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8765))
    app.run(host='0.0.0.0', port=port, debug=False)
