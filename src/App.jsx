import React, { useState, useCallback } from 'react';
import { PDFDocument } from 'pdf-lib';
import * as pdfjsLib from 'pdfjs-dist';
import JSZip from 'jszip';
import { Document, Packer, Paragraph, ImageRun } from 'docx';

// Initialize PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

function App() {
  const [extractedImages, setExtractedImages] = useState([]);
  const [loading, setLoading] = useState(false);
  const [pageCount, setPageCount] = useState(0);
  const [startPage, setStartPage] = useState(1);
  const [endPage, setEndPage] = useState(1);
  const [imagesPerPage, setImagesPerPage] = useState(1);

  const extractImagesFromPDF = async (file, start, end) => {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    setPageCount(pdf.numPages);
    
    const images = [];
    for (let i = start; i <= Math.min(end, pdf.numPages); i++) {
      const page = await pdf.getPage(i);
      const ops = await page.getOperatorList();
      const imgs = ops.fnArray.reduce((acc, fn, idx) => {
        if (fn === pdfjsLib.OPS.paintImageXObject) {
          acc.push(ops.argsArray[idx][0]);
        }
        return acc;
      }, []);

      for (const imgName of imgs) {
        try {
          const img = await page.objs.get(imgName);
          const canvas = document.createElement('canvas');
          canvas.width = img.width;
          canvas.height = img.height;
          const ctx = canvas.getContext('2d');
          ctx.putImageData(new ImageData(img.data, img.width, img.height), 0, 0);
          
          images.push({
            data: canvas.toDataURL(),
            page: i,
            originalSize: [img.width, img.height]
          });
        } catch (error) {
          console.error('Error extracting image:', error);
        }
      }
    }
    return images;
  };

  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    setLoading(true);
    try {
      const images = await extractImagesFromPDF(file, startPage, endPage);
      setExtractedImages(images);
    } catch (error) {
      console.error('Error processing file:', error);
      alert('Error processing file. Please try again.');
    } finally {
      setLoading(false);
    }
  };

  const exportToPDF = async () => {
    const pdfDoc = await PDFDocument.create();
    const imagesPerRow = Math.ceil(Math.sqrt(Number(imagesPerPage)));
    const rows = Math.ceil(imagesPerPage / imagesPerRow);

    for (let i = 0; i < extractedImages.length; i += Number(imagesPerPage)) {
      const page = pdfDoc.addPage();
      const { width, height } = page.getSize();
      const margin = 50;
      const availableWidth = width - (margin * 2);
      const availableHeight = height - (margin * 2);
      const imageWidth = availableWidth / imagesPerRow;
      const imageHeight = availableHeight / rows;

      const pageImages = extractedImages.slice(i, i + Number(imagesPerPage));
      for (let j = 0; j < pageImages.length; j++) {
        const row = Math.floor(j / imagesPerRow);
        const col = j % imagesPerRow;
        const imageData = await fetch(pageImages[j].data);
        const imageBytes = await imageData.arrayBuffer();
        const image = await pdfDoc.embedPng(imageBytes);
        
        page.drawImage(image, {
          x: margin + (col * imageWidth),
          y: height - margin - ((row + 1) * imageHeight),
          width: imageWidth - 10,
          height: imageHeight - 10,
        });
      }
    }

    const pdfBytes = await pdfDoc.save();
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'extracted_images.pdf';
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportToZip = async () => {
    const zip = new JSZip();
    extractedImages.forEach((image, index) => {
      const imageData = image.data.split(',')[1];
      zip.file(`image_${index + 1}_page${image.page}.png`, imageData, { base64: true });
    });

    const content = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(content);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'extracted_images.zip';
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportToDocx = async () => {
    const doc = new Document({
      sections: [{
        properties: {},
        children: await Promise.all(
          extractedImages.map(async (image, index) => {
            const response = await fetch(image.data);
            const blob = await response.blob();
            const arrayBuffer = await blob.arrayBuffer();
            
            return new Paragraph({
              children: [
                new ImageRun({
                  data: arrayBuffer,
                  transformation: {
                    width: 500,
                    height: 300
                  }
                })
              ]
            });
          })
        )
      }]
    });

    const buffer = await Packer.toBlob(doc);
    const url = URL.createObjectURL(buffer);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'extracted_images.docx';
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="bg-gray-100 min-h-screen p-8">
      <div className="max-w-4xl mx-auto">
        <h1 className="text-3xl font-bold mb-8">Document Image Extractor</h1>
        
        <div className="bg-white rounded-lg shadow p-6 mb-8">
          <form className="space-y-4" onSubmit={(e) => e.preventDefault()}>
            <div>
              <label className="block text-sm font-medium text-gray-700">Upload PDF Document</label>
              <input
                type="file"
                accept=".pdf"
                onChange={handleFileUpload}
                className="mt-1 block w-full border-gray-300 rounded-md shadow-sm"
              />
            </div>
            
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700">Start Page</label>
                <input
                  type="number"
                  min="1"
                  max={pageCount}
                  value={startPage}
                  onChange={(e) => setStartPage(Number(e.target.value))}
                  className="mt-1 block w-full border-gray-300 rounded-md shadow-sm"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700">End Page</label>
                <input
                  type="number"
                  min={startPage}
                  max={pageCount}
                  value={endPage}
                  onChange={(e) => setEndPage(Number(e.target.value))}
                  className="mt-1 block w-full border-gray-300 rounded-md shadow-sm"
                />
              </div>
            </div>
          </form>
        </div>

        {loading && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
            <div className="bg-white p-6 rounded-lg shadow-xl text-center">
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
              <p className="text-lg font-semibold">Processing Document...</p>
              <p className="text-sm text-gray-600">Please wait while we extract the images.</p>
            </div>
          </div>
        )}

        {extractedImages.length > 0 && (
          <div className="bg-white rounded-lg shadow p-6">
            <h2 className="text-xl font-semibold mb-4">Extracted Images</h2>
            <div className="grid grid-cols-2 md:grid-cols-3 gap-4">
              {extractedImages.map((image, index) => (
                <div key={index} className="relative aspect-square">
                  <img
                    src={image.data}
                    alt={`Extracted image ${index + 1}`}
                    className="w-full h-full object-contain border rounded"
                  />
                  <div className="absolute top-2 right-2 bg-blue-600 text-white px-2 py-1 rounded text-sm">
                    Page {image.page}
                  </div>
                </div>
              ))}
            </div>

            <div className="mt-6 pt-6 border-t">
              <h3 className="text-lg font-medium mb-3">Export Options</h3>
              <div className="space-y-3">
                <div>
                  <label className="block text-sm font-medium text-gray-700">Images per page</label>
                  <input
                    type="number"
                    min="1"
                    max="8"
                    value={imagesPerPage}
                    onChange={(e) => setImagesPerPage(Number(e.target.value))}
                    className="mt-1 block w-full border-gray-300 rounded-md shadow-sm"
                  />
                </div>
                <div className="flex flex-col sm:flex-row space-y-2 sm:space-y-0 sm:space-x-4">
                  <button
                    onClick={exportToPDF}
                    className="flex-1 bg-green-600 text-white px-4 py-2 rounded-md hover:bg-green-700"
                  >
                    Export to PDF
                  </button>
                  <button
                    onClick={exportToDocx}
                    className="flex-1 bg-red-600 text-white px-4 py-2 rounded-md hover:bg-red-700"
                  >
                    Export to Word
                  </button>
                  <button
                    onClick={exportToZip}
                    className="flex-1 bg-yellow-600 text-white px-4 py-2 rounded-md hover:bg-yellow-700"
                  >
                    Download as ZIP
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default App;