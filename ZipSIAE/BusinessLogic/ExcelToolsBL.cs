using System;
using System.Collections.Generic;
using System.Linq;


using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Drawing;
using System.Drawing.Imaging;

using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.IO;
using System.Drawing.Drawing2D;

namespace BusinessLogic
{
 public   class ExcelToolsBL
    {

        //INSTRUCCION 1
        private static Row GetRow(Worksheet worksheet, int rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
                    Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        //INSTRUCCION 2
        private static Cell GetCell(Worksheet worksheet, String columName, int rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);
            if (row == null) return null;

            return row.Elements<Cell>().Where(c => String.Compare
                    (c.CellReference.Value, columName + rowIndex, true) == 0).First();

        }

        //INSTRUCCION 3
        private static WorksheetPart GetWorkSheetPartByName(SpreadsheetDocument document, String sheetName)
        {

            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                //the worksheet especificado no existe
                return null;
            }
            String relationShipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
                        document.WorkbookPart.GetPartById(relationShipId);
            return worksheetPart;

        }

        //INSTRUCCION 4
        public static void UpdateCell(String rutaDest, String nameSheet, String dato, int rowIndex, String columName)
        {

            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(rutaDest, true))
            {
                WorksheetPart worksheetPart = GetWorkSheetPartByName(spreadSheet, nameSheet);

                if (worksheetPart != null)
                {

                    Cell cell = GetCell(worksheetPart.Worksheet, columName, rowIndex);
                    cell.CellValue = new CellValue(dato);
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);

                    //guardar el workSheet
                    worksheetPart.Worksheet.Save();

                }

            }

        }


        public static ImagePartType GetImagePartTypeByBitmap(Bitmap image)
        {
            if (ImageFormat.Bmp.Equals(image.RawFormat))
                return ImagePartType.Bmp;
            else if (ImageFormat.Gif.Equals(image.RawFormat))
                return ImagePartType.Gif;
            else if (ImageFormat.Png.Equals(image.RawFormat))
                return ImagePartType.Png;
            else if (ImageFormat.Tiff.Equals(image.RawFormat))
                return ImagePartType.Tiff;
            else if (ImageFormat.Icon.Equals(image.RawFormat))
                return ImagePartType.Icon;
            else if (ImageFormat.Jpeg.Equals(image.RawFormat))
                return ImagePartType.Jpeg;
            else if (ImageFormat.Emf.Equals(image.RawFormat))
                return ImagePartType.Emf;
            else if (ImageFormat.Wmf.Equals(image.RawFormat))
                return ImagePartType.Wmf;
            else
                throw new Exception("Image type could not be determined.");
        }

        public static void AddImage(WorksheetPart worksheetPart,
                                Stream imageStream, string imgDesc,
                                int colNumber, int rowNumber, int width, int height)
        {
            // Necesitamos la transmisión de imágenes más de una vez, así creamos una copia de memoria
             MemoryStream imageMemStream = new MemoryStream();
            

            imageStream.Position = 0;
            imageStream.CopyTo(imageMemStream);
            imageStream.Position = 0;



            var drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null)
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

            if (!worksheetPart.Worksheet.ChildElements.OfType<Drawing>().Any())
            {
                worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            if (drawingsPart.WorksheetDrawing == null)
            {
                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            }

            var worksheetDrawing = drawingsPart.WorksheetDrawing;



            // Bitmap bm = new Bitmap(imageMemStream);
          Bitmap bm = new Bitmap(ConvertImage(imageMemStream,ImageFormat.Jpeg));

            bm.SetResolution(60, 60);

            var imagePart = drawingsPart.AddImagePart(GetImagePartTypeByBitmap(bm));
           //imagePart.FeedData(imageStream);
            imagePart.FeedData(ConvertImage(imageStream,ImageFormat.Jpeg));
            
           
            A.Extents extents = new A.Extents();
            // var extentsCx = bm.Width * (long)(914400 / bm.HorizontalResolution);
            //var extentsCy = bm.Height * (long)(914400 / bm.VerticalResolution);
            //dividir  las medidas hecnas en el paint por 1.35 esa era la medida para 
            //el programa

            //var extentsCx = width * (long)(914400 / bm.HorizontalResolution);
            //var extentsCy = height * (long)(914400 / bm.VerticalResolution);

            var extentsCx = width * (long)(571500 / bm.HorizontalResolution);
            var extentsCy = height * (long)(571500 / bm.VerticalResolution);



            bm.Dispose();

            var colOffset = 0;
            var rowOffset = 0;

            var nvps = worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>();
            var nvpId = nvps.Count() > 0
                ? (UInt32Value)worksheetDrawing.Descendants<Xdr.NonVisualDrawingProperties>().Max(p => p.Id.Value) + 1
                : 1U;

            var oneCellAnchor = new Xdr.OneCellAnchor(
                new Xdr.FromMarker
                {
                    ColumnId = new Xdr.ColumnId((colNumber - 1).ToString()),
                    RowId = new Xdr.RowId((rowNumber - 1).ToString()),
                    ColumnOffset = new Xdr.ColumnOffset(colOffset.ToString()),
                    RowOffset = new Xdr.RowOffset(rowOffset.ToString())
                },
                new Xdr.Extent { Cx = extentsCx, Cy = extentsCy },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = nvpId, Name = "Picture " + nvpId, Description = imgDesc },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })
                    ),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = drawingsPart.GetIdOfPart(imagePart), CompressionState = A.BlipCompressionValues.Print },
                        new A.Stretch(new A.FillRectangle())
                    ),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0, Y = 0 },
                            new A.Extents { Cx = extentsCx, Cy = extentsCy }
                        ),
                        new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle }
                    )
                ),
                new Xdr.ClientData()
            );

            worksheetDrawing.Append(oneCellAnchor);
        }


        public static void AddImageDocument(bool createFile, string excelFile, string sheetName,
                                   Stream imageStream, string imgDesc,
                                   int rowNumber, int colNumber, int width, int height)
        {
            SpreadsheetDocument spreadsheetDocument = null;
            WorksheetPart worksheetPart = null;
            if (createFile)
            {
                // Create a spreadsheet document by supplying the filepath
                spreadsheetDocument = SpreadsheetDocument.Create(excelFile, SpreadsheetDocumentType.Workbook);

                // Add a WorkbookPart to the document
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart
                worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                };
                sheets.Append(sheet);
            }
            else
            {
                // Open spreadsheet
                spreadsheetDocument = SpreadsheetDocument.Open(excelFile, true);

                // Get WorksheetPart
                worksheetPart = GetWorkSheetPartByName(spreadsheetDocument, sheetName);
            }

            AddImage(worksheetPart, imageStream, imgDesc, colNumber, rowNumber, width, height);

            worksheetPart.Worksheet.Save();

            spreadsheetDocument.Close();
        }

        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

        public static Stream ConvertImage( Stream originalStream, ImageFormat format)
        {
            var image = Image.FromStream(originalStream);
            var stream = new MemoryStream();
            image.Save(stream, format);
            stream.Position = 0;
            return stream;
        }

    }


}
