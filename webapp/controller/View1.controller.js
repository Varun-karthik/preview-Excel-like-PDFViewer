sap.ui.define([
    "sap/ui/core/mvc/Controller",
    'sap/m/MessageToast'
], (Controller,MessageToast) => {
    "use strict";

    return Controller.extend("excelpreview.controller.View1", {
        onInit() {
            this.filedata = null;
            this._oHot = null;
        },
        handleUploadComplete: function(oEvent) {
            var oFile = oEvent.getParameter("files") ? oEvent.getParameter("files")[0] : oEvent;
            var reader = new FileReader();
            var that = this;
            var container = document.getElementById('hot-container');

            reader.onload = function (e) {
                // 1. Read the file data
                var data = new Uint8Array(e.target.result);
                var workbook = XLSX.read(data, { type: 'array' });
                
                // 2. Get the first sheet
                var sSheetName = workbook.SheetNames[0];
                
                // 3. Convert sheet to Array of Arrays (Required for Handsontable)
                var aJsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sSheetName], {
                    header: 1,  // Forces an array format instead of objects
                    defval: ""  // Prevents empty cells from collapsing
                });

                that.filedata = aJsonData;
                console.log("Parsed Excel Data:", that.filedata);

                // 4. Initialize or update Handsontable
                if (that._oHot) {
                    // If table already exists, just load the new data
                    that._oHot.loadData(that.filedata);
                } else {
                    // Create the table for the first time
                    that._oHot = new Handsontable(container, {
                        data: that.filedata,
                        rowHeaders: true,
                        colHeaders: true,
                        height: 'auto',
                        autoWrapRow: true,
                        autoWrapCol: true,
                        licenseKey: 'non-commercial-and-evaluation', 
                        contextMenu: true,
                        filters: true,
                        dropdownMenu: true,
                        manualColumnResize: true,
                        manualRowResize: true,
                        copyPaste: true
                    });
                }
            };

            // Start reading the file
            if (oFile) {
                reader.readAsArrayBuffer(oFile);
            }
		},

		handleUploadPress: function() {
            var oFileUploader = this.byId("fileUploader");
            oFileUploader.checkFileReadable().then(function() {
                oFileUploader.upload();
                MessageToast.show("File uploaded successfully.");
            }, function(error) {
                MessageToast.show("The file cannot be read. It may have changed.");
            });
		},

        handlePreviewPress: function() {
            var oFileUploader = this.byId("fileUploader");
            // Access the underlying native HTML input element from the UI5 wrapper
            var oDomRef = oFileUploader.getDomRef();
            var oFileInput = oDomRef ? oDomRef.querySelector("input[type='file']") : null;
            var oFile = oFileInput && oFileInput.files ? oFileInput.files[0] : null;
            if (!oFile) {
                MessageToast.show("Please select a valid Excel file first.");
                return;
            }

            var reader = new FileReader();
            var that = this;
            
            // Locate the container injected via the core:HTML control in our view
            var container = document.getElementById('hot-container');

            if (!container) {
                MessageToast.show("Preview container element not found.");
                return;
            }

            reader.onload = function(e) {
                try {
                    // 1. Read the raw buffer data
                    var data = new Uint8Array(e.target.result);
                    var workbook = XLSX.read(data, { type: 'array' });
                    
                    // 2. Fetch data out of the first workbook tab
                    var sSheetName = workbook.SheetNames[0];
                    
                    // 3. Convert target sheet into an Array of Arrays
                    var aJsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sSheetName], {
                        header: 1,  // Array format response
                        defval: ""  // Protects from cell-collapse anomalies
                    });

                    that.filedata = aJsonData;
                    console.log("Parsed Excel Data for Preview:", that.filedata);

                    // 4. Instantiate or update the Handsontable matrix display
                    if (that._oHot) {
                        that._oHot.loadData(that.filedata);
                    } else {
                        that._oHot = new Handsontable(container, {
                            data: that.filedata,
                            rowHeaders: true,
                            colHeaders: true,
                            width: '100%',
                            height: 'auto',
                            autoWrapRow: true,
                            autoWrapCol: true,
                            licenseKey: 'non-commercial-and-evaluation', 
                            contextMenu: true,
                            filters: true,
                            dropdownMenu: true,
                            manualColumnResize: true,
                            manualRowResize: true,
                            copyPaste: true
                        });
                    }
                } catch (oError) {
                    console.error(oError);
                    MessageToast.show("Failed to parse and preview Excel document structure.");
                }
            };

            // Initialize reading buffer pipeline
            reader.readAsArrayBuffer(oFile);
        }
    });
});