# DocsToPdfConverter

[![License](https://img.shields.io/badge/license-MIT%20License-blue.svg)](LICENSE)

Converts docx, xls, xlsx, pptx into PDF

For docx conversion, [docx4j-export-FO](https://github.com/plutext/docx4j-export-FO) is needed

Note for xls and xlsx to pdf conversion, any charts included in the spreadsheet will not showup in pdf (So far not supported). Suggestions are welcomed

## Known problems

- Cannot render Chinese characters not in [Arial, Times New Roman, Helvetica, Calibri, 等线, 宋体] 
- High resolution images in slideshow will not be rendered
- Background image in slideshow may not be properly rendered

## Dependencies 

Note: The version numbers listed are the newest ones when I am making the project

- [Docx4J 3.3.5](https://www.docx4java.org)

- ANTLR 2.2.7 and ANTLR Runtime 3.5.2

- Apache Avalon 4.3.1

- Apache Batik 1.9

- Apache Commons Codec 1.10

- Apache Commons Collections4 4.1

- Apache Commons Fop 2.2

- Apache Commons IO 2.4

- Apache Commons Lang3 3.4

- Apache Commons Logging 1.2

- Apache Fontbox 2.0.4

- [Apache POI 3.16](https://poi.apache.org)

- [Apache XMLBeans 2.6.0](https://mvnrepository.com/artifact/org.apache.xmlbeans/xmlbeans/2.6.0)

- Apache Xerces XML Serializer 2.7.2

- Apache XML Graphics Commons 2.2

- [FlyingSaucer API](https://code.google.com/archive/p/flying-saucer/)

- [Google Guava 19.0](https://github.com/google/guava)

- JAXB-XSL-FO 1.0.1

- Slf4J 1.4.21

- Xalan 2.7.1

