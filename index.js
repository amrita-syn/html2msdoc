const html2word = (html, options = {}) => {
  const defaultOptions = {
    filename: "document.docx",
    margins: {
      ...{ top: 1, bottom: 1, left: 1, right: 1, header: 0.25, footer: 0.25 },
      ...options.margins,
    },
    styles: "",
    headerHTML: "",
    headerStyle: "",
    orientation: "portrait",
    documentProperties: {},
  };
  delete options.margins;
  const {
    filename,
    margins,
    styles,
    headerHTML,
    headerStyle,
    orientation,
    documentProperties,
  } = {
    ...defaultOptions,
    ...options,
  };

  const wordXML = `
  <html 
    xmlns:v='urn:schemas-microsoft-com:vml'
    xmlns:o='urn:schemas-microsoft-com:office:office'
    xmlns:w='urn:schemas-microsoft-com:office:word'
    xmlns:m='http://schemas.microsoft.com/office/2004/12/omml'
    xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
      <meta http-equiv=Content-Type content="text/html; charset=utf-8">
      <title>${filename}</title>
      <style>
        v\:* { behavior:url(#default#VML); }
        o\:* { behavior:url(#default#VML); }
        w\:* { behavior:url(#default#VML); }
        .shape { behavior:url(#default#VML); }
      
        @page
        {
          mso-page-orientation: ${orientation};
          size: 8.27in 11.69in;
          margin: ${margins.top}in ${margins.right}in ${margins.bottom}in ${
    margins.left
  }in;
        }
        @page Section1 {
          mso-header-margin:${margins.header}in;
          mso-footer-margin:${margins.footer}in;
          mso-header: h1;
          mso-footer: f1;
        }
        div.Section1 { page:Section1; }
        div#h1
        {
          margin:0in 0in 0in 900in;
          width:1px;
          height:1px;
          overflow:hidden;
        }
        table {
          border-collapse: collapse;
        }
        table td {
          border: 1pt solid gray;
          padding: 0.05in 0.1in;
        }
        ${headerStyle}
        ${styles}
      </style>
      <xml>
        <w:WordDocument>
          <w:View>Print</w:View>
          <w:Zoom>100</w:Zoom>
          <w:DoNotOptimizeForBrowser/>
        </w:WordDocument>
        <o:DocumentProperties>
          <o:Title>${documentProperties["Title"] || ""}</o:Title>
          <o:Author>${documentProperties["Author"] || ""}</o:Author>
          <o:Created>${documentProperties["Created"] || ""}</o:Created>
          <o:Company>${documentProperties["Company"] || ""}</o:Company>
          <o:Version>${documentProperties["Version"] || ""}</o:Version>
        </o:DocumentProperties>
      </xml>
    </head>
      
    <body>
      <div class="Section1">
      
        ${html}
    
        <div style='mso-element:header' id=h1 >
          <!-- HEADER-tags -->
          ${headerHTML}
          <!-- end HEADER-tags -->
        </div>
      </div>
    </body>
  </html>>`;
  const source =
    "data:application/vnd.ms-word;charset=utf-8," + encodeURIComponent(wordXML);
  const fileDownload = document.createElement("a");
  document.body.appendChild(fileDownload);
  fileDownload.href = source;
  fileDownload.download = "document.doc";
  fileDownload.click();
  document.body.removeChild(fileDownload);
};



export default { html2word };
