using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocWD = DocumentFormat.OpenXml.Wordprocessing;
using DVML = DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.IO;
using System.Configuration;

namespace Web02.Utilities
{
    public class Utilities
    {
        public MemoryStream GenerarActa()
        {
            MemoryStream stream = new MemoryStream();
            using (WordprocessingDocument doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                // Agregar un cuerpo al documento
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new DocWD.Document();
                DocWD.Body body = mainPart.Document.AppendChild(new DocWD.Body());

                // Configurar el tamaño de página y los márgenes
                DocWD.SectionProperties sectionProps = new DocWD.SectionProperties(
                    new DocWD.PageMargin()
                    {
                        Left = 1700,
                        Right = 1700,
                        Top = 2270,
                        Bottom = 1700
                    }, new DocWD.PageSize() { Width = 11906, Height = 16838, Orient = DocWD.PageOrientationValues.Portrait }
                );
                body.AppendChild(sectionProps);

                #region HEADER - FOOTER

                AgregarEncabezadoDocumento(mainPart);
                AgregarPiePaginaDocumentoDigea(mainPart);

                #endregion

                DocWD.Paragraph subtitleParagraph = body.AppendChild(new DocWD.Paragraph(new DocWD.ParagraphProperties(new DocWD.Justification() { Val = DocWD.JustificationValues.Center })));
                DocWD.Run runSubtitle = new DocWD.Run(new DocWD.Text("ACTA DE VISITA DE PROTECCIÓN DE DERECHOS"));
                runSubtitle.RunProperties = new DocWD.RunProperties(new DocWD.Bold(), new DocWD.FontSize() { Val = "22" });
                subtitleParagraph.AppendChild(runSubtitle);

                AgregarSaltoDeLinea(body);

                DocWD.TableProperties tablePropertiesHeader = ObtenerPropiedadesTabla();
                DocWD.Table tableHeader = new DocWD.Table();
                tableHeader.Append(tablePropertiesHeader);
                AgregarFilaTablaEtiquetaInformeDigea(tableHeader, "NOMBRE DE IPRESS", ":", $"a1");
                AgregarFilaTablaEtiquetaInformeDigea(tableHeader, "CÓDIGO ÚNICO", ":", $"a2");
                AgregarFilaTablaEtiquetaInformeDigea(tableHeader, "DIRECCIÓN", ":", $"a3 -DEPARTAMENTO DE a4 - PROVINCIA DE a5 - DISTRITO DE a6 ");
                AgregarFilaTablaEtiquetaInformeDigea(tableHeader, "TURNO SUPERVISADO", ":", "- ");
                body.AppendChild(tableHeader);

                AgregarSaltoDeLinea(body);

                //OPERARIOS
                AgregarTextoSimple(body, $"Siendo las *****, se hicieron presentes el equipo supervisor del FISSAL, conformado por:", false, 22);

                AgregarTextoSimple(body, $"Participantes de la IPRESS a7:", true, 22);

                AgregarTextoSimple(body, $"Se adjuntan los siguientes documentos los cuales forman parte del acta de vista de protección de derechos:", false, 22);
                AgregarTextoSimple(body, $"Instrumento de protección de derechos", false, 22);
                AgregarSaltoDeLinea(body);
                AgregarTextoSimple(body, $"Se realizó la visita encontrándose lo siguiente:", true);
                AgregarTextoSimple(body, $"De un total de a8 pacientes entrevistados:", false);

                AgregarSaltoDeLinea(body);

                AgregarTextoSimple(body, $"La IPRESS se compromete a enviar la información solicitada el día de hoy a9 del presente, hasta las 17:15 horas.", false, 22);
                AgregarSaltoDeLinea(body);
                AgregarTextoSimple(body, $"Siendo las a10, firman los presentes en señal de conformidad de los hallazgos encontrados y registrados en el anexo adjunto a la presente acta.", false, 22);
                AgregarSaltoDeLinea(body);

                AgregarTextoSimple(body, $"Firmas:", false, 22);

                AgregarSaltoDePagina(body);

                //RESPUESTAS IPRESS
                AgregarTextoSimpleCentrado(body, $"INSTRUMENTO DE PROTECCIÓN DE DERECHOS - IPRESS", true, 22);

                AgregarSaltoDePagina(body);

                //RESPUESTAS ASEGURADOS
                AgregarTextoSimpleCentrado(body, $"INSTRUMENTO DE PROTECCIÓN DE DERECHOS - ASEGURADOS", true, 22);
                AgregarSaltoDeLinea(body);
                AgregarTextoSimple(body, $"- NA: No aplica.", false, 22);
                mainPart.Document.Save();
            }

            return stream;
        }

        public void AgregarTextoSimpleCentrado(DocWD.Body body, string text, bool bold, int size = 22)
        {
            DocWD.ParagraphProperties paragraphProperties = new DocWD.ParagraphProperties(new DocWD.Justification() { Val = DocWD.JustificationValues.Center });
            DocWD.Paragraph paragraph = body.AppendChild(new DocWD.Paragraph(paragraphProperties));
            DocWD.Run run = new DocWD.Run(new DocWD.Text(text));
            run.RunProperties = new DocWD.RunProperties(new DocWD.FontSize() { Val = size.ToString() });
            if (bold)
            {
                run.RunProperties.Bold = new DocWD.Bold();
            }
            paragraph.AppendChild(run);
        }

        public void AgregarSaltoDePagina(DocWD.Body body)
        {
            DocWD.Paragraph paragraph = body.AppendChild(new DocWD.Paragraph());
            paragraph.AppendChild(new DocWD.Run(new DocWD.Break() { Type = DocWD.BreakValues.Page }));
            body.AppendChild(new DocWD.Paragraph());
        }

        public void AgregarTextoSimple(DocWD.Body body, string text, bool bold, int size = 22)
        {
            DocWD.ParagraphProperties paragraphProperties = new DocWD.ParagraphProperties(new DocWD.Justification() { Val = DocWD.JustificationValues.Both });
            DocWD.Paragraph paragraph = body.AppendChild(new DocWD.Paragraph(paragraphProperties));
            DocWD.Run run = new DocWD.Run(new DocWD.Text(text));
            run.RunProperties = new DocWD.RunProperties(new DocWD.FontSize() { Val = size.ToString() });
            if (bold)
            {
                run.RunProperties.Bold = new DocWD.Bold();
            }
            paragraph.AppendChild(run);
        }

        public void AgregarFilaTablaEtiquetaInformeDigea(DocWD.Table tabla, string etiqueta, string interceptor, string valor)
        {
            DocWD.TableRow fila = new DocWD.TableRow();

            DocWD.TableCell celdaEtiqueta = new DocWD.TableCell();
            DocWD.TableProperties tableProperties2 = new DocWD.TableProperties(
                new DocWD.TableCellMarginDefault(
                    new DocWD.LeftMargin() { Width = "40" },
                    new DocWD.RightMargin() { Width = "40" }
                )
            );
            tabla.AppendChild(tableProperties2);

            DocWD.Paragraph etiquetaParagraph = new DocWD.Paragraph();
            DocWD.Run runEtiqueta = new DocWD.Run(new DocWD.Text(etiqueta));
            runEtiqueta.RunProperties = new DocWD.RunProperties(new DocWD.Bold(), new DocWD.FontSize() { Val = "22" });
            etiquetaParagraph.AppendChild(runEtiqueta);
            etiquetaParagraph.ParagraphProperties = new DocWD.ParagraphProperties(new DocWD.Justification() { Val = DocWD.JustificationValues.Both });
            celdaEtiqueta.Append(etiquetaParagraph);

            DocWD.TableCellMarginDefault cellMarginDefault = new DocWD.TableCellMarginDefault();
            DocWD.LeftMargin leftMargin = new DocWD.LeftMargin() { Width = "40" };
            DocWD.RightMargin rightMargin = new DocWD.RightMargin() { Width = "40" };
            cellMarginDefault.Append(leftMargin);
            cellMarginDefault.Append(rightMargin);
            celdaEtiqueta.Append(cellMarginDefault);
            fila.AppendChild(celdaEtiqueta);

            DocWD.TableCell celdaInterceptor = new DocWD.TableCell();
            DocWD.Paragraph interceptorParagraph = new DocWD.Paragraph();
            DocWD.Run runInterceptor = new DocWD.Run(new DocWD.Text(interceptor));
            runInterceptor.RunProperties = new DocWD.RunProperties(new DocWD.FontSize() { Val = "22" });
            interceptorParagraph.AppendChild(runInterceptor);
            interceptorParagraph.ParagraphProperties = new DocWD.ParagraphProperties(new DocWD.Justification() { Val = DocWD.JustificationValues.Both });
            celdaInterceptor.Append(interceptorParagraph);
            fila.AppendChild(celdaInterceptor);

            DocWD.TableCell celdaValor = new DocWD.TableCell();
            DocWD.Paragraph valorParagraph = new DocWD.Paragraph();
            DocWD.Run runValor = new DocWD.Run(new DocWD.Text(valor));
            runValor.RunProperties = new DocWD.RunProperties(new DocWD.FontSize() { Val = "22" });
            valorParagraph.AppendChild(runValor);
            valorParagraph.ParagraphProperties = new DocWD.ParagraphProperties(new DocWD.Justification() { Val = DocWD.JustificationValues.Both });
            celdaValor.Append(valorParagraph);
            fila.AppendChild(celdaValor);


            tabla.AppendChild(fila);

            if (tabla.ChildElements.Count == 1)
            {
                DocWD.TableProperties tableProperties = new DocWD.TableProperties();
                DocWD.TableWidth tableWidth = new DocWD.TableWidth() { Width = "2000", Type = DocWD.TableWidthUnitValues.Pct };
                tableProperties.Append(tableWidth);
                tabla.AppendChild(tableProperties);
            }
        }

        public DocWD.TableProperties ObtenerPropiedadesTabla()
        {
            DocWD.TableProperties tableProperties = new DocWD.TableProperties();
            DocWD.TableWidth tableWidth = new DocWD.TableWidth() { Width = "5000", Type = DocWD.TableWidthUnitValues.Pct };
            tableProperties.Append(tableWidth);
            return tableProperties;
        }

        public void AgregarSaltoDeLinea(DocWD.Body body)
        {
            DocWD.ParagraphProperties paragraphProps = new DocWD.ParagraphProperties();
            DocWD.SpacingBetweenLines spacing = new DocWD.SpacingBetweenLines() { After = "0" };
            paragraphProps.Append(spacing);
            DocWD.Paragraph paragraph = new DocWD.Paragraph(paragraphProps);
            body.AppendChild(paragraph);
        }

        public void GenerarPiePaginaDigea(FooterPart footerPart)
        {
            DocWD.Footer footer = new DocWD.Footer();
            DocWD.Paragraph paragraph1 = new DocWD.Paragraph();
            DocWD.Run run1 = new DocWD.Run();
            DocWD.Picture picture1 = new DocWD.Picture();
            DVML.Shape shape1 = new DVML.Shape()
            {
                Id = "WordPictureWatermarkF01",
                Style = $"left:0;text-align:center;margin-left:0;margin-top:50.0pt;width:{LeerVariablesConfig("WidthActa")};height:{LeerVariablesConfig("HeightActa")};",
                OptionalString = "_x0000_s2051",
                AllowInCell = false,
                Type = "#_x0000_t75"
            };
            DVML.ImageData imageData1 = new DVML.ImageData() { Title = "Footer", RelationshipId = "rId000" };
            shape1.Append(imageData1);
            picture1.Append(shape1);
            run1.Append(picture1);
            paragraph1.Append(run1);

            footer.Append(paragraph1);
            footerPart.Footer = footer;
        }

        public string ObtenerImagenB64(string file)
        {
            FileStream inFile;
            byte[] byteArray;
            try
            {
                inFile = new FileStream(file, FileMode.Open, FileAccess.Read);
                byteArray = new byte[inFile.Length];
                long byteRead = inFile.Read(byteArray, 0, (int)inFile.Length);
                inFile.Close();
                return Convert.ToBase64String(byteArray, 0, byteArray.Length);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void GenerarEncabezado(HeaderPart headerPart)
        {
            DocWD.Header header = new DocWD.Header();

            DocWD.ParagraphProperties paragraphProps = new DocWD.ParagraphProperties();
            DocWD.SpacingBetweenLines spacing = new DocWD.SpacingBetweenLines() { After = "0" };
            paragraphProps.Append(spacing);

            DocWD.Paragraph paragraph1 = new DocWD.Paragraph();
            DocWD.Paragraph paragraph2 = new DocWD.Paragraph(new DocWD.ParagraphProperties(new DocWD.Justification() { Val = DocWD.JustificationValues.Center }, paragraphProps));
            DocWD.Paragraph paragraph3 = new DocWD.Paragraph(new DocWD.ParagraphProperties(new DocWD.Justification() { Val = DocWD.JustificationValues.Center }));

            DocWD.Run run1 = new DocWD.Run();
            DocWD.Picture picture1 = new DocWD.Picture();
            DVML.Shape shape1 = new DVML.Shape()
            {
                Id = "WordPictureWatermarkH01",
                Style = $"left:0;text-align:center;margin-left:0;margin-top:0;width:{LeerVariablesConfig("WidthActa")};height:{LeerVariablesConfig("HeightActa")};mso-position-horizontal:right;",
                OptionalString = "_x0000_s2051",
                AllowInCell = false,
                Type = "#_x0000_t75"
            };
            DVML.ImageData imageData1 = new DVML.ImageData() { Title = "Header", RelationshipId = "rId999" };
            shape1.Append(imageData1);
            picture1.Append(shape1);
            run1.Append(picture1);
            paragraph1.Append(run1);

            DocWD.Run run2 = new DocWD.Run(new DocWD.Text("“" + LeerVariablesConfig("TextPrincipalActa") + "”"));
            run2.RunProperties = new DocWD.RunProperties(new DocWD.FontSize() { Val = "13" });
            paragraph2.Append(run2);
            DocWD.Run run3 = new DocWD.Run(new DocWD.Text("“" + LeerVariablesConfig("TextSecondaryActa") + "”"));
            run3.RunProperties = new DocWD.RunProperties(new DocWD.FontSize() { Val = "13" });
            paragraph3.Append(run3);

            header.Append(paragraph1);
            header.Append(paragraph2);
            header.Append(paragraph3);
            headerPart.Header = header;
        }

        public void GenerarImagenEncabezadoPie(ImagePart imagePart1, string imageData)
        {
            using (Stream data = new MemoryStream(Convert.FromBase64String(imageData)))
            {
                imagePart1.FeedData(data);
            }
        }

        public void AgregarEncabezadoDocumento(MainDocumentPart mainPart)
        {
            string ImageB64 = ObtenerImagenB64(LeerVariablesConfig("UrlActa"));
            HeaderPart headPart1 = mainPart.AddNewPart<HeaderPart>();
            GenerarEncabezado(headPart1);
            string rId = mainPart.GetIdOfPart(headPart1);
            ImagePart image = headPart1.AddNewPart<ImagePart>("image/jpeg", "rId999");
            GenerarImagenEncabezadoPie(image, ImageB64);
            IEnumerable<DocWD.SectionProperties> sectPrs = mainPart.Document.Body.Elements<DocWD.SectionProperties>();
            foreach (var sectPr in sectPrs)
            {
                sectPr.RemoveAllChildren<DocWD.HeaderReference>();
                sectPr.PrependChild(new DocWD.HeaderReference() { Id = rId });
            }
        }

        public void AgregarPiePaginaDocumentoDigea(MainDocumentPart mainPart)
        {
            string ImageB64Footer = ObtenerImagenB64(LeerVariablesConfig("UrlActa"));
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();
            GenerarPiePaginaDigea(footerPart);
            string footerRId = mainPart.GetIdOfPart(footerPart);
            ImagePart imageFooter = footerPart.AddNewPart<ImagePart>("image/jpeg", "rId000");
            GenerarImagenEncabezadoPie(imageFooter, ImageB64Footer);
            IEnumerable<DocWD.SectionProperties> sectPrsFooter = mainPart.Document.Body.Elements<DocWD.SectionProperties>();
            foreach (var sectPr in sectPrsFooter)
            {
                sectPr.RemoveAllChildren<DocWD.FooterReference>();
                sectPr.PrependChild(new DocWD.FooterReference() { Id = footerRId });

                var pageMargin = sectPr.Descendants<DocWD.PageMargin>().FirstOrDefault();
                if (pageMargin != null)
                {
                    pageMargin.Bottom = 2550;
                }
            }
        }

        public static string LeerVariablesConfig(string variable)
        {
            return ConfigurationManager.AppSettings[variable];
        }
    }
}