const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '../web')));

const upload = multer({ dest: path.join(__dirname, '../uploads/') });

function parseExcelData(jsonData) {
    if (!jsonData || jsonData.length < 1) {
        throw new Error('Excel文件数据为空');
    }

    let headerRowIndex = 0;
    let headers = [];
    
    for (let i = 0; i < Math.min(5, jsonData.length); i++) {
        const row = jsonData[i];
        if (!row) continue;
        
        const rowHeaders = row.map(h => String(h || '').trim());
        const hasName = rowHeaders.some(h => h.includes('姓名') || h.includes('名字'));
        
        if (hasName) {
            headerRowIndex = i;
            headers = rowHeaders;
            break;
        }
    }

    if (headers.length === 0) {
        throw new Error('未找到表头行，请确保Excel包含姓名列');
    }

    const nameIndex = headers.findIndex(h => h.includes('姓名') || h.includes('名字'));
    const classIndex = headers.findIndex(h => h.includes('班级') || h.includes('班') || h.includes('课程') || h.includes('班级名称'));
    const amountIndex = headers.findIndex(h => h.includes('金额') || h.includes('费用') || h.includes('钱') || h.includes('退费') || h.includes('应收'));
    const timeIndex = headers.findIndex(h => h.includes('最后修改时间') || h.includes('修改时间') || h.includes('退费时间') || h.includes('时间') || h.includes('收费时间'));

    if (nameIndex === -1) {
        throw new Error('未找到姓名列，请确保Excel包含姓名列');
    }

    const data = [];
    const startRow = headerRowIndex + 1;
    
    for (let i = startRow; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || !row[nameIndex]) continue;

        const name = String(row[nameIndex]).trim();
        if (!name) continue;

        data.push({
            name: name,
            className: classIndex >= 0 ? String(row[classIndex] || '').trim() : '',
            amount: amountIndex >= 0 ? parseAmount(row[amountIndex]) : 0,
            refundTime: timeIndex >= 0 ? formatTime(row[timeIndex]) : ''
        });
    }

    return data;
}

function formatTime(value) {
    if (!value) return '';
    
    if (typeof value === 'number') {
        const excelDate = value < 60 ? value - 1 : value - 2;
        const jsDate = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
        return jsDate.toLocaleString('zh-CN');
    }
    
    const str = String(value).trim();
    if (str.includes('T')) {
        return str.replace('T', ' ').substring(0, 19);
    }
    
    return str;
}

function parseAmount(value) {
    if (typeof value === 'number') return value;
    const str = String(value).replace(/[^0-9.]/g, '');
    return parseFloat(str) || 0;
}

function processRefundData(paymentData, refundData) {
    const paymentNameCounts = {};
    const refundNameCounts = {};
    const refundDetails = {};

    paymentData.forEach(item => {
        paymentNameCounts[item.name] = (paymentNameCounts[item.name] || 0) + 1;
    });

    refundData.forEach(item => {
        refundNameCounts[item.name] = (refundNameCounts[item.name] || 0) + 1;
        
        if (!refundDetails[item.name]) {
            refundDetails[item.name] = [];
        }
        refundDetails[item.name].push({
            className: item.className,
            amount: item.amount,
            refundTime: item.refundTime
        });
    });

    const pureRefund = [];
    Object.keys(refundNameCounts).forEach(name => {
        if (paymentNameCounts[name] === refundNameCounts[name]) {
            const details = refundDetails[name];
            details.forEach(detail => {
                pureRefund.push({
                    name: name,
                    className: detail.className,
                    amount: detail.amount,
                    refundTime: detail.refundTime
                });
            });
        }
    });

    return pureRefund;
}

app.post('/api/refund/process', upload.fields([{ name: 'paymentFile' }, { name: 'refundFile' }]), (req, res) => {
    try {
        if (!req.files || !req.files.paymentFile || !req.files.refundFile) {
            return res.status(400).json({ error: '请上传缴费和退费两个Excel文件' });
        }

        const paymentFilePath = req.files.paymentFile[0].path;
        const refundFilePath = req.files.refundFile[0].path;

        const paymentWorkbook = XLSX.readFile(paymentFilePath);
        const refundWorkbook = XLSX.readFile(refundFilePath);

        const paymentWorksheet = paymentWorkbook.Sheets[paymentWorkbook.SheetNames[0]];
        const refundWorksheet = refundWorkbook.Sheets[refundWorkbook.SheetNames[0]];

        const paymentJson = XLSX.utils.sheet_to_json(paymentWorksheet, { header: 1 });
        const refundJson = XLSX.utils.sheet_to_json(refundWorksheet, { header: 1 });

        const paymentData = parseExcelData(paymentJson);
        const refundData = parseExcelData(refundJson);

        const result = processRefundData(paymentData, refundData);

        fs.unlinkSync(paymentFilePath);
        fs.unlinkSync(refundFilePath);

        const totalAmount = result.reduce((sum, item) => sum + item.amount, 0);
        const uniqueNames = new Set(result.map(item => item.name)).size;

        res.json({
            success: true,
            data: result,
            totalCount: result.length,
            uniqueCount: uniqueNames,
            totalAmount: totalAmount
        });
    } catch (error) {
        if (req.files && req.files.paymentFile) {
            try { fs.unlinkSync(req.files.paymentFile[0].path); } catch (e) {}
        }
        if (req.files && req.files.refundFile) {
            try { fs.unlinkSync(req.files.refundFile[0].path); } catch (e) {}
        }
        res.status(500).json({ error: error.message });
    }
});

app.get('/api/refund/health', (req, res) => {
    res.json({ status: 'ok', message: '退费系统API正常运行' });
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, '../web/refund.html'));
});

app.listen(PORT, () => {
    console.log(`纯退费系统运行在端口 ${PORT}`);
    console.log(`访问地址: http://localhost:${PORT}`);
});