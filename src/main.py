import os
import zipfile
import xml
from xml.etree import ElementTree


'''
book.xlsx/
 |- _rels/
 |- xl/
 |  |- drawings/
 |  |  |  |- _rels/a
 |  |  |  |- drawing1.xml.rels # 画像にRIDを割り当て
 |  |  |- drawing1.xml         # 画像の位置やサイズなどの情報
 |  |-media/
 |  | |- image1.png            # 画像ファイル
 |  | |- image2.jpg
 |  |- workbook.xml            # Bookの情報
'''

def emus_to_pt(emus: int) -> int:
    return emus / 12700

class Relationship:
    """
    * §9.2
    * *.rels (workbook.xml.rels) files
    * Assign a rid to a target like a sheet or a image
    """
    @staticmethod
    def from_xml(el: xml.etree.ElementTree.Element):
        return Relationship(
            el.get('Id'),
            el.get('Target'),
            el.get('Type'),
        )

    def __init__(self, id: str, target: str, type: str):
        self.id = id
        self.target = target
        self.type = type

    def __str__(self) -> str:
        return str(self.__dict__)

class Relationships:
    """
    * §9.2
    """
    @staticmethod
    def from_archive(xlsx: zipfile.ZipFile, path: str):
        rels_xml_str = xlsx.read(path).decode()
        rels_el = ElementTree.fromstring(rels_xml_str)
        ns = {
            '': 'http://schemas.openxmlformats.org/package/2006/relationships',
        }
        relationships = rels_el.findall('.//Relationship', ns)
        items = [];
        for r in relationships:
            items.append(Relationship.from_xml(r))
        return Relationships(items)

    def __init__(self, items: list[Relationship]):
        self.items = items

    def __getitem__(self, index):
        return self.items[index]

    def __setitem__(self, index):
        return self.items[index]

    def __len__(self):
        return len(self.items)

    def get(self, rid: str) -> Relationship:
        for i in self.items:
            if i.id == rid:
                return i
        return None



class AnchorPoint:
    """
    * §20.5.2.15 (from) or v20.5.2.32 (to)
    * xdr:from or xdr:to
    """
    def __init__(self, col: int, row: int):
        self.col = col
        self.row = row

class TwoCellAnchor:
    @staticmethod
    def from_xml(el: xml.etree.ElementTree, rels: Relationships, file_path: str):
        dir_path = os.path.dirname(file_path)

        ns = {
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        }
        fromPt = AnchorPoint(
            col=int(el.find('.//xdr:from/xdr:col', ns).text),
            row=int(el.find('.//xdr:from/xdr:row', ns).text),
        )
        toPt = AnchorPoint(
            col=int(el.find('.//xdr:to/xdr:col', ns).text),
            row=int(el.find('.//xdr:to/xdr:row', ns).text),
        )
        title = el.find('.//xdr:pic/xdr:nvPicPr/xdr:cNvPr', ns).get('name')
        cx = el.find('.//xdr:pic/xdr:spPr/a:xfrm/a:ext', ns).get('cx')
        cy = el.find('.//xdr:pic/xdr:spPr/a:xfrm/a:ext', ns).get('cy')
        embed_rid = el.find('.//*/xdr:blipFill/a:blip', ns).get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        embed = rels.get(embed_rid)
        image_path = os.path.normpath(os.path.join(dir_path, embed.target))
        return TwoCellAnchor(fromPt, toPt, title, cx, cy, embed, image_path)

    def __init__(
            self,
            fromPt: AnchorPoint,
            toPt: AnchorPoint,
            title: str,
            cx: int, cy: int,
            embed: Relationship,
            image_path: str):
        self.fromPt = fromPt
        self.toPt = toPt
        self.title = title
        self.embed = embed
        self.image_path = image_path
        self.cy = cy
        self.cx = cx

    def __str__(self) -> str:
        return f'from:{self.fromPt.col},{self.fromPt.row}, to:{self.toPt.col},{self.toPt.row}, embed:{self.embed}'

class SpreadsheetDrawing:
    @staticmethod
    def from_archive(xlsx: zipfile.ZipFile, path: str):
        dir_path = os.path.dirname(path)
        file_name = os.path.basename(path)
        drawing_xml_str = xlsx.read(path).decode()
        rels_path = os.path.join(dir_path, '_rels', f'{file_name}.rels')

        drawing_xml_rels = Relationships.from_archive(xlsx, rels_path)

        drawing_el = ElementTree.fromstring(drawing_xml_str)
        ns = {
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        }
        two_cell_anchors = drawing_el.findall('.//xdr:twoCellAnchor', ns)
        items = []
        for anchor in two_cell_anchors:
            items.append(TwoCellAnchor.from_xml(anchor, drawing_xml_rels, path))
        return SpreadsheetDrawing(items)

    def __init__(self, two_cell_anchors: list[TwoCellAnchor]):
        self.two_cell_anchors = two_cell_anchors

class Sheet:
    @staticmethod
    def from_archive(xlsx: zipfile.ZipFile, path: str, name: str, id: str):
        dir_path = os.path.dirname(path)
        file_name = os.path.basename(path)
        sheet_xml_str = xlsx.read(path).decode()
        rels_path = os.path.join(dir_path, '_rels', f'{file_name}.rels')

        sheet_xml_rels = Relationships.from_archive(xlsx, rels_path)

        ns = {
            '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        }
        sheet_xml_el = ElementTree.fromstring(sheet_xml_str)
        drawing_el = sheet_xml_el.find('.//drawing', ns)
        if drawing_el != None:
            rid = drawing_el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            rel = sheet_xml_rels.get(rid)
            if rel != None:
                drawing_path = os.path.normpath(os.path.join(dir_path, rel.target))
                drawing = SpreadsheetDrawing.from_archive(xlsx, drawing_path)

        return Sheet(name, id, drawing)

    def __init__(self, name: str, id: str, drawing: SpreadsheetDrawing):
        self.name = name
        self.id = id
        self.drawing = drawing

class Workbook:
    @staticmethod
    def parse(path: str):
        xlsx = zipfile.ZipFile(xlsx_path)

        workbook_xml_str = xlsx.read('xl/workbook.xml').decode()

        workbook_rels = Relationships.from_archive(xlsx, 'xl/_rels/workbook.xml.rels')

        ns = {
            '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        }
        workbook_el = ElementTree.fromstring(workbook_xml_str)
        sheet_els = workbook_el.findall('.//sheets/sheet', ns)

        sheets = []
        for sheet_el in sheet_els:
            name = sheet_el.get('name')
            id = sheet_el.get('sheetid')
            rid = sheet_el.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            rel = workbook_rels.get(rid);
            sheets.append(Sheet.from_archive(xlsx, f'xl/{rel.target}', name, id))

        return Workbook(sheets)

    def __init__(self, sheets: list[Sheet]):
        self.sheets = sheets

# Load book
xlsx_path = os.path.realpath(os.path.join(os.path.dirname(__file__), '../data/Book1.xlsx'))
xlsx = zipfile.ZipFile(xlsx_path)

# Workbook
print('\n# Workbook:\n')
wb = Workbook.parse(xlsx_path)
for sheet in wb.sheets:
    print('Sheet:', sheet.name)
    if sheet.drawing != None:
        for anchor in sheet.drawing.two_cell_anchors:
            print('  ', anchor.toPt.col, anchor.toPt.row, anchor.image_path)
