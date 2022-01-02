window = this;
document = { createElementNS: function () { return {}; } };


importScripts('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.66/pdfmake.min.js');
importScripts('https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.71/vfs_fonts.min.js');

onmessage = function(req) {
  new Promise(function (resolve, reject) {
    generatePdfBlob(req.data, function (result) {
      if (result) { resolve(result); } else { reject(); }
    });
  }).then(function (blob) {
    postMessage({ blob });
  });
};
function generatePdfBlob(data, callback) {
  if (!callback) {
    throw new Error('generatePdfBlob is an async method and needs a callback');
  }
  const docDefinition = generateDocDefinition(data);
  pdfMake.createPdf(docDefinition).getBlob(callback);
}
function generateDocDefinition(data) {
  return data;
}