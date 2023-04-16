## xdr:colOffの単位

セル内のオフセット量を示す。単位はEMUs。

* 20.5.2.11 colOff (Column Offset) で、ST_Coordinateで定義されるとある
* 20.1.10.16 ST_Coordinate (Coordinate) で定義される
    * EMUs (English Metric Units)
    * ST_CoordinateUnqualifiedとST_UniversalMeasureのUnion
        * ST_UniversalMeasureで1 ptが1/72 inchと定義される
* 20.1.10.35 ST_LineWidthに1 ptが12700 EMUsと定義される
* colOff/12700で簡易的なPixel値となり、colOff/12700/72でインチになる

## リンク

* [ECMA-376 Office Open XML file formats](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
    * Part 1のPDFファイル (Fundamentals And Markup Language Reference) が詳しい
