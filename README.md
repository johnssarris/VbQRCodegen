<div align="center">
<img width="418" height="463" src="https://dl.unicontsoft.com/upload/pix/ss_qr_code5.png">

## VbQRCodegen
QR Code generator library for VBA (MS Office)
</div>

### Description

A single file QR Code generator based on https://www.nayuki.io/page/qr-code-generator-library

### Requirements

- MS Office with VBA (Excel, Access, Word, etc.)
- Windows only (uses Windows GDI/OLE APIs)
- Ensure **OLE Automation** is checked under Tools → References in the VBA editor

### Usage

Add `mdQRCodegen.bas` to your VBA project and call `QRCodegenBarcode` to get a picture from text or a byte array:

```vba
    Set Image1.Picture = QRCodegenBarcode("Sample text")
```
The returned picture uses vectors (Enhanced Metafile), so it scales to any size without loss of quality.

### MS Access Support

For compatibility with image controls on forms/reports you can use `QRCodegenConvertToData` function like this:
```
    Image0.PictureData = QRCodegenConvertToData(QRCodegenBarcode("Sample text"))
```
If this does not work in your version of MS Access (for some reason) then you can try converting QR Code to a bitmap instead like this:
```
    Image0.PictureData = QRCodegenConvertToData(QRCodegenBarcode("Sample text"), 500, 500)
```
Note that this produces 500x500 bitmap picture of the QR Code so might need to tweak output size parameters.
