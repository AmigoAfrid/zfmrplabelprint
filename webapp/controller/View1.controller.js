sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/m/MessageBox",
    "zfmrplabelprint/util/xlsx",
    "sap/ui/export/Spreadsheet",
    "zfmrplabelprint/util/PDFLib",


], (Controller, MessageBox, xlsx, Spreadsheet) => {
    "use strict";

    return Controller.extend("zfmrplabelprint.controller.View1", {
        onInit() {

            this._pdfViewer = new sap.m.PDFViewer({
                isTrustedSource: true,
                width: "100%",
                height: "600px", // Adjust height as needed
                title: ""
            });
            this.getView().addDependent(this._pdfViewer);

        },

        onmrplabelprintUploadChange: function (oEvent) {
            this._file = oEvent.getParameter("files")?.[0] || null;
        },

        // onmrplabelprintUploadPress: function () {

        //     if (!this._file) {
        //         sap.m.MessageToast.show("Please select a file first");
        //         return;
        //     }

        //     const reader = new FileReader();

        //     reader.onload = (e) => {

        //         try {
        //             const data = new Uint8Array(e.target.result);

        //             const workbook = XLSX.read(data, { type: "array" });

        //             let tableData = [];

        //             workbook.SheetNames.forEach(sheetName => {

        //                 const rows = XLSX.utils.sheet_to_json(
        //                     workbook.Sheets[sheetName]
        //                 );

        //                 rows.forEach(row => {
        //                     tableData.push({
        //                         Material            : row.Material || "",
        //                         Description         : row.Description || "",
        //                         Quantity            : row.Quantity || 0,
        //                         NoLabel             : row["No Label"] || 0,
        //                         LabelQuantity       : row["Label Quantity"] || 0,
        //                         DateOfManufacture   : row["Date Of Manufacture"] || ""
        //                     });
        //                 });
        //             });

        //             const oModel = new sap.ui.model.json.JSONModel({
        //                 Datass: tableData
        //             });

        //             this.getView().setModel(oModel, "ColorUploadModel");

        //             sap.m.MessageToast.show("Excel Data Loaded Successfully");

        //         } catch (err) {
        //             sap.m.MessageBox.error("Failed to read Excel file.");
        //             console.error(err);
        //         }
        //     };

        //     reader.readAsArrayBuffer(this._file);
        // },
        onmrplabelprintUploadPress: function () {

            if (!this._file) {
                sap.m.MessageToast.show("Please select a file first");
                return;
            }

            const reader = new FileReader();

            reader.onload = (e) => {

                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: "array" });

                    let ProductEntries = [];

                    // Excel Data
                    workbook.SheetNames.forEach(sheetName => {
                        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

                        rows.forEach(row => {
                            if (row.Product) {
                                ProductEntries.push({
                                    Product: row.Product,
                                    Quantity: row.Quantity || 0,
                                    DateOfManufacture: row.DateOfManufacture || ""
                                });
                            }
                        });
                    });

                    if (ProductEntries.length === 0) {
                        sap.m.MessageToast.show("No Products found in Excel");
                        return;
                    }

                    const oODataModel = this.getView().getModel("ZSB_MRPLABEL_DESC");
                    this.getView().setBusy(true);

                    // Fetching Product Description
                    const promises = ProductEntries.map((oEntry) => {

                        return new Promise((resolve) => {

                            oODataModel.read("/ZMRP_PRODUCTDESCRIPTION", {
                                filters: [
                                    new sap.ui.model.Filter(
                                        "Product",
                                        sap.ui.model.FilterOperator.EQ,
                                        oEntry.Product
                                    )
                                ],
                                success: function (oData) {

                                    resolve({
                                        Product: oEntry.Product,
                                        Quantity: oEntry.Quantity,
                                        DateOfManufacture: oEntry.DateOfManufacture,
                                        ProductDescription: oData.results?.[0]?.ProductDescription || ""
                                    });

                                },
                                error: function () {
                                    resolve({
                                        Product: oEntry.Product,
                                        Quantity: oEntry.Quantity,
                                        DateOfManufacture: oEntry.DateOfManufacture,
                                        ProductDescription: ""
                                    });
                                }
                            });

                        });

                    });

                    // 🔹 Set Model
                    Promise.all(promises).then((tableData) => {

                        const oJsonModel = new sap.ui.model.json.JSONModel({
                            Datass: tableData
                        });

                        this.getView().setModel(oJsonModel, "ColorUploadModel");

                        this.getView().setBusy(false);

                        sap.m.MessageToast.show("Data fetched successfully");

                    }).catch(() => {
                        this.getView().setBusy(false);
                        sap.m.MessageBox.error("Error while fetching data");
                    });

                } catch (err) {
                    this.getView().setBusy(false);
                    sap.m.MessageBox.error("Failed to process file.");
                    console.error(err);
                }
            };

            reader.readAsArrayBuffer(this._file);
        },
        onDownloadTemplate: function () {
            var that = this;

            MessageBox.confirm("Do you want to download the template?", {
                onClose: function (oAction) {
                    if (oAction === MessageBox.Action.OK) {

                        var aTemplateData = that.createTemplateData();
                        var aCols = that.createColumnsCodeGroup();

                        var oSettings = {
                            workbook: {
                                columns: aCols
                            },
                            dataSource: aTemplateData,
                            fileName: "MRP_Label_Template.xlsx"
                        };

                        var oSheet = new sap.ui.export.Spreadsheet(oSettings);
                        oSheet.build().finally(function () {
                            oSheet.destroy();
                        });
                    }
                }
            });
        },

        createColumnsCodeGroup: function () {
            return [
                { label: "Product", property: "Name", type: "string" },
                { label: "Quantity", property: "Quantity", type: "number" },
                { label: "DateOfManufacture", property: "DateOfManufacture", type: "date", format: "yyyy-MM-dd" }
            ];
        },

        createTemplateData: function () {
            return [
                {
                    Name: "",
                    ProductId: "",
                    Quantity: 0,
                    NoLabel: 0,
                    LabelQuantity: 0,
                    DateOfManufacture: ""
                }
            ];
        },

        onPrint: async function () {
            sap.ui.core.BusyIndicator.show();

            var oTable = this.byId("mrptable");
            var aSelectedIndices = oTable.getSelectedIndices();

            if (aSelectedIndices.length === 0) {
                sap.m.MessageToast.show("Please select at least one row");
                sap.ui.core.BusyIndicator.hide();
                return;
            }

            var oModel = this.getView().getModel("ColorUploadModel");
            var aData = oModel.getProperty("/Datass");

            try {
                const pdfContentArray = [];


                for (let i = 0; i < aSelectedIndices.length; i++) {
                    const iIndex = aSelectedIndices[i];
                    const oRow = aData[iIndex];
                    const oMaterial = oRow.Product;
                    const oQty = oRow.Quantity;
                    const oDOM = String(oRow.DateOfManufacture);

                    const sServiceUrl = `/sap/bc/http/sap/ZMRP_LABEL_PRINT?product=${oMaterial}&Labelquantity=${oQty}&Dateofmanufacture=${oDOM}`;
                    console.log("Service URL:", sServiceUrl);

                    const pdfData = await this.fetchPDFData(sServiceUrl);

                    pdfContentArray.push(pdfData);
                }

                // Display all PDFs
                this.displayPDFs(pdfContentArray);

                sap.m.MessageToast.show("PDF(s) displayed successfully!");

            } catch (error) {
                console.error("Error generating PDF:", error);
                sap.m.MessageToast.show("Failed to generate PDF. Please try again.");
            } finally {
                sap.ui.core.BusyIndicator.hide();
            }
        },
        fetchPDFData: function (sServiceUrl) {
            return new Promise((resolve, reject) => {
                jQuery.ajax({
                    url: sServiceUrl,
                    method: "GET",
                    success: function (data, textStatus, jqXHR) {
                        resolve(data); // Resolve with PDF data
                    },
                    error: function (jqXHR, textStatus, errorThrown) {
                        console.error("Error fetching data. Status:", textStatus, "Error:", errorThrown);
                        sap.m.MessageToast.show("HTTP Service Error...!");
                        reject(errorThrown);
                        sap.ui.core.BusyIndicator.hide();
                    }
                });
            });
        },

        displayPDFs: async function (pdfDataArray) {
            const base64ToArrayBuffer = (base64) => {
                const binaryString = atob(base64);
                const binaryLen = binaryString.length;
                const bytes = new Uint8Array(binaryLen);
                for (let i = 0; i < binaryLen; i++) {
                    bytes[i] = binaryString.charCodeAt(i);
                }
                return bytes.buffer;
            };

            const mergedPdf = await PDFLib.PDFDocument.create();
            for (let document of pdfDataArray) {
                const pdfBytes = base64ToArrayBuffer(document);
                const pdfDoc = await PDFLib.PDFDocument.load(pdfBytes);
                const copiedPages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices());
                copiedPages.forEach((page) => mergedPdf.addPage(page));
            }

            const pdfBytes = await mergedPdf.save();
            const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
            const _pdfurl = URL.createObjectURL(pdfBlob);

            if (!this._pdfViewer) {
                this._pdfViewer = new sap.m.PDFViewer({
                    width: "auto",
                    source: _pdfurl
                });
                jQuery.sap.addUrlWhitelist("blob");
            } else {
                this._pdfViewer.setSource(_pdfurl);
            }

            this._pdfViewer.setTitle("MRP Label");
            this._pdfViewer.open();
        },
        onResetPress: function () {
            var oModel = this.getView().getModel("ColorUploadModel");
            oModel.setProperty("/Datass", []);
        }

    });
});