import React, { useState, useRef } from 'react';
import { Upload, Calendar, Download, FileSpreadsheet, FileText, AlertCircle, RefreshCw } from 'lucide-react';
import './App.css';

const ListingCalculator = () => {
  const [excelFile, setExcelFile] = useState(null);
  const [pdfFiles, setPdfFiles] = useState([]);
  const [categoryMapFile, setCategoryMapFile] = useState(null);
  const [postageRateFile, setPostageRateFile] = useState(null);
  const [date, setDate] = useState('');
  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [progress, setProgress] = useState('');
  const [extractedData, setExtractedData] = useState(null);
  const [pdfDataArray, setPdfDataArray] = useState([]);
  const workerRef = useRef(null);
  
  const excelInputRef = useRef(null);
  const pdfInputRef = useRef(null);
  const categoryInputRef = useRef(null);
  const postageInputRef = useRef(null);

  const formatDateToSKU = (dateStr) => {
    const [year, month, day] = dateStr.split('-');
    return `${day}${month}${year.slice(2)}`;
  };

  const clearAllInputs = () => {
    setExcelFile(null);
    setPdfFiles([]);
    setCategoryMapFile(null);
    setPostageRateFile(null);
    setDate('');
    setError('');
    setSuccess('');
    setProgress('');
    setExtractedData(null);
    setPdfDataArray([]);
    
    if (excelInputRef.current) excelInputRef.current.value = '';
    if (pdfInputRef.current) pdfInputRef.current.value = '';
    if (categoryInputRef.current) categoryInputRef.current.value = '';
    if (postageInputRef.current) postageInputRef.current.value = '';
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      setExcelFile(file);
      setError('');
    }
  };

  const handlePdfUpload = (e) => {
    const files = Array.from(e.target.files);
    if (files.length > 0) {
      setPdfFiles(files);
      setError('');
    }
  };

  const handleCategoryMapUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      setCategoryMapFile(file);
      setError('');
    }
  };

  const handlePostageRateUpload = (e) => {
    const file = e.target.files[0];
    if (file) {
      setPostageRateFile(file);
      setError('');
    }
  };

  const processFiles = async () => {
    if (!excelFile || pdfFiles.length === 0 || !date) {
      setError('Please upload Excel file, at least one PDF, and select a date');
      return;
    }

    setProcessing(true);
    setError('');
    setSuccess('');
    setProgress('Loading PDF.js...');
    setExtractedData(null);

    try {
      if (!window.pdfjsLib) {
        await new Promise((resolve, reject) => {
          const script = document.createElement('script');
          script.src = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js';
          script.onload = resolve;
          script.onerror = reject;
          document.head.appendChild(script);
        });
        
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = 
          'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
      }

      setProgress('Extracting text from PDFs...');
      const pdfDataArray = [];
      
      for (let pdfIndex = 0; pdfIndex < pdfFiles.length; pdfIndex++) {
        const pdfFile = pdfFiles[pdfIndex];
        setProgress(`Processing PDF ${pdfIndex + 1} of ${pdfFiles.length}...`);
        
        const pdfArrayBuffer = await pdfFile.arrayBuffer();
        const loadingTask = window.pdfjsLib.getDocument({ data: pdfArrayBuffer });
        const pdfDoc = await loadingTask.promise;
        
        let pdfText = '';
        
        for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
          const page = await pdfDoc.getPage(pageNum);
          const textContent = await page.getTextContent();
          const pageText = textContent.items.map(item => item.str).join(' ');
          pdfText += pageText + '\n';
        }
        
        const shippingMatch = pdfText.match(/Net\s+shipping.*?£\s*(\d+\.?\d*)/i) || 
                            pdfText.match(/shipping.*?£\s*(\d+\.?\d*)/i);
        const shipping = shippingMatch ? parseFloat(shippingMatch[1]) : 0;
        
        const invoiceDateMatch = pdfText.match(/ORDER\s+DATE[:\s]+(\d{2}[-/]\d{2}[-/]\d{4})/i);
        const invoiceDate = invoiceDateMatch ? invoiceDateMatch[1] : '';
        
        const vendorNumberMatch = pdfText.match(/INVOICE\s+([A-Z0-9]+)/i);
        const vendorNumber = vendorNumberMatch ? vendorNumberMatch[1] : '';
        
        pdfDataArray.push({
          text: pdfText,
          shipping: shipping,
          invoiceDate: invoiceDate,
          vendorNumber: vendorNumber,
          name: pdfFile.name,
          index: pdfIndex
        });
      }

      setPdfDataArray(pdfDataArray);

      setProgress('Initializing worker...');
      if (!workerRef.current) {
        workerRef.current = new Worker(
          new URL('./workers/fileProcessor.js', import.meta.url),
          { type: 'classic' }
        );
      }

      const worker = workerRef.current;

      worker.onmessage = (e) => {
        const { type, message, data, url, filename, csvUrl, csvFilename, error: workerError } = e.data;

        if (type === 'progress') {
          setProgress(message);
        } else if (type === 'extracted') {
          setExtractedData(data);
        } else if (type === 'success') {
          setProgress('Downloading files...');
          
          const a = document.createElement('a');
          a.href = url;
          a.download = filename;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);

          const csvLink = document.createElement('a');
          csvLink.href = csvUrl;
          csvLink.download = csvFilename;
          document.body.appendChild(csvLink);
          csvLink.click();
          document.body.removeChild(csvLink);
          URL.revokeObjectURL(csvUrl);

          setSuccess(message);
          setProgress('');
          setProcessing(false);
          
          setTimeout(() => {
            clearAllInputs();
          }, 3000);
        } else if (type === 'error') {
          setError('Error: ' + workerError);
          setProgress('');
          setProcessing(false);
        }
      };

      worker.onerror = (error) => {
        setError('Worker error: ' + error.message);
        setProgress('');
        setProcessing(false);
      };

      const excelBuffer = await excelFile.arrayBuffer();

      let categoryBuffer = null;
      if (categoryMapFile) {
        categoryBuffer = await categoryMapFile.arrayBuffer();
      }

      let postageBuffer = null;
      if (postageRateFile) {
        postageBuffer = await postageRateFile.arrayBuffer();
      }

      worker.postMessage({
        excelBuffer,
        pdfDataArray,
        categoryBuffer,
        postageBuffer,
        date
      });

    } catch (err) {
      setError('Error: ' + err.message);
      setProcessing(false);
      setProgress('');
      console.error(err);
    }
  };

  return (
    <div className="app-container">
      <div className="background-overlay"></div>
      
      <div className="main-wrapper">
        <header className="app-header">
          <div className="header-content">
            <div className="header-left">
              <h1 className="app-title">S3G SYNERGY</h1>
              <p className="app-subtitle">Listing Calculator</p>
            </div>
            <div className="status-indicator">
              <div className="status-dot"></div>
              <span>System Active</span>
            </div>
          </div>
        </header>

        <main className="main-content">
          <div className="content-card">
            <p className="content-subtitle">
              Generate Excel listing + eBay CSV from invoice & product data
            </p>

            {error && (
              <div className="alert alert-error">
                <AlertCircle className="alert-icon" />
                <p>{error}</p>
              </div>
            )}

            {success && (
              <div className="alert alert-success">
                <p className="alert-title">{success}</p>
                <p className="alert-detail">✨ Inputs will auto-clear in 3 seconds...</p>
                <p className="alert-detail">📊 Column A: Order Date | Column B: Vendor | Column E: Manifest SKU</p>
              </div>
            )}

            {extractedData && (
              <div className="alert alert-info">
                <h4 className="info-title">📄 Extracted from PDF:</h4>
                <div className="info-content">
                  <p><strong>Total Shipping:</strong> £{extractedData.totalShipping.toFixed(2)}</p>
                  <p><strong>SKUs Found:</strong> {extractedData.items.filter(i => i.cost > 0).length} of {extractedData.items.length}</p>
                  {pdfDataArray && pdfDataArray.length > 0 && (
                    <>
                      <p className="invoice-header"><strong>📅 Invoice Info:</strong></p>
                      {pdfDataArray.map((pdf, idx) => (
                        <p key={idx} className="invoice-detail">
                          PDF {idx + 1}: {pdf.invoiceDate || 'N/A'} | Vendor: {pdf.vendorNumber || 'N/A'}
                        </p>
                      ))}
                    </>
                  )}
                </div>
              </div>
            )}

            <div className="form-grid">
              <div className="form-card">
                <div className="form-header">
                  <div className="form-icon form-icon-purple">
                    <Calendar size={20} />
                  </div>
                  <h3>Date for SKU</h3>
                </div>
                <input
                  type="date"
                  value={date}
                  onChange={(e) => setDate(e.target.value)}
                  className="date-input"
                />
                {date && (
                  <p className="sku-preview">
                    SKU prefix: <span className="sku-code">{formatDateToSKU(date)}</span>
                  </p>
                )}
              </div>

              <div className="form-card">
                <div className="form-header">
                  <div className="form-icon form-icon-purple">
                    <FileSpreadsheet size={20} />
                  </div>
                  <h3>Category Map</h3>
                </div>
                <input
                  ref={categoryInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleCategoryMapUpload}
                  className="file-input-hidden"
                  id="category-upload"
                />
                <label htmlFor="category-upload" className="upload-button upload-button-purple">
                  <Upload size={16} />
                  <span>{categoryMapFile ? categoryMapFile.name : 'Upload Excel'}</span>
                </label>
              </div>

              <div className="form-card">
                <div className="form-header">
                  <div className="form-icon form-icon-orange">
                    <FileSpreadsheet size={20} />
                  </div>
                  <h3>Postage Rate</h3>
                </div>
                <input
                  ref={postageInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handlePostageRateUpload}
                  className="file-input-hidden"
                  id="postage-upload"
                />
                <label htmlFor="postage-upload" className="upload-button upload-button-orange">
                  <Upload size={16} />
                  <span>{postageRateFile ? postageRateFile.name : 'Upload Excel'}</span>
                </label>
              </div>

              <div className="form-card">
                <div className="form-header">
                  <div className="form-icon form-icon-green">
                    <FileSpreadsheet size={20} />
                  </div>
                  <h3>Excel File</h3>
                </div>
                <input
                  ref={excelInputRef}
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleExcelUpload}
                  className="file-input-hidden"
                  id="excel-upload"
                />
                <label htmlFor="excel-upload" className="upload-button upload-button-green">
                  <Upload size={16} />
                  <span>{excelFile ? excelFile.name : 'Choose Excel'}</span>
                </label>
              </div>
            </div>

            <div className="pdf-upload-section">
              <div className="form-header">
                <div className="form-icon form-icon-red">
                  <FileText size={20} />
                </div>
                <h3>PDF Invoice(s)</h3>
              </div>
              <input
                ref={pdfInputRef}
                type="file"
                accept=".pdf"
                onChange={handlePdfUpload}
                className="file-input-hidden"
                id="pdf-upload"
                multiple
              />
              <label htmlFor="pdf-upload" className="upload-button upload-button-red">
                <Upload size={16} />
                <span>
                  {pdfFiles.length > 0 
                    ? `${pdfFiles.length} PDF${pdfFiles.length > 1 ? 's' : ''} selected` 
                    : 'Choose PDF(s)'}
                </span>
              </label>
              {pdfFiles.length > 0 && (
                <div className="file-list">
                  {pdfFiles.map((file, idx) => (
                    <div key={idx} className="file-item">
                      <div className="file-dot"></div>
                      <span>{file.name}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>

            <button
              onClick={processFiles}
              disabled={!excelFile || pdfFiles.length === 0 || !date || processing}
              className="generate-button"
            >
              {processing ? (
                <>
                  <div className="spinner"></div>
                  Processing...
                </>
              ) : (
                <>
                  <Download size={20} />
                  Generate Excel + eBay CSV
                </>
              )}
            </button>

            {progress && (
              <div className="progress-container">
                <div className="progress-content">
                  <div className="progress-spinner"></div>
                  <div className="progress-info">
                    <p>{progress}</p>
                    <div className="progress-bar">
                      <div className="progress-bar-fill"></div>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {!processing && !success && (
              <div className="footer-actions">
                <div className="info-box">
                  <h4>✨ Optimized Performance</h4>
                  <p>Background processing - UI stays smooth!</p>
                </div>
                
                {(excelFile || pdfFiles.length > 0 || categoryMapFile || postageRateFile || date) && (
                  <button onClick={clearAllInputs} className="clear-button">
                    <RefreshCw size={16} />
                    Clear All
                  </button>
                )}
              </div>
            )}
          </div>
        </main>

        <footer className="app-footer">
          <p>© 2025 S3G Synergy. All rights reserved.</p>
        </footer>
      </div>
    </div>
  );
};

export default ListingCalculator;