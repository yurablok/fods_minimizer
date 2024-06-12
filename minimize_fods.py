#!/usr/bin/env python3
#
# Flat XML ODF Spreadsheets Minimizer
#
# Cuts out all unnecessary and optimizes styles in spreadsheet documents
# to make it much easier to solve conflicts of merging in version control systems.
#
# Launch it in the directory that contains `.fods` documents.
#
# - Deduplication, sorting, and stable names of styles (based on their hash).
# - Cuts out all styles except the color of cells text and the color of cells background.
# - Preserves frozen columns.
# - Reduces documents size by ~60%.
#
# Author: Yurii Blok
# License: BSL-1.0
# https://github.com/yurablok/fods_minimizer
# History:
# v1.0 2024-Jun-12     First release.

import os, io
import xml.etree.ElementTree as ET
import xml.dom.minidom
import xml.sax.saxutils
import hashlib


for _, _, fileNames in os.walk("."):
    for fileName in fileNames:
        if not fileName.endswith(".fods") or fileName.startswith("_"):
            continue
        print("-" * len(fileName))
        print(fileName)
        print("-" * len(fileName))

        buffer = io.StringIO()
        doc = xml.sax.saxutils.XMLGenerator(buffer, "UTF-8", True)

        stack = [True] * 64
        #stack.append(True)
        stackIdx = 0
        stackMax = 0
        countStart = 0
        countChars = 0
        countEnd = 0


        class MyContentHandler(xml.sax.ContentHandler):
            def __init__(self):
                self.isOfficeStyles = False
                self.styles = {}
                self.styleName = ""
                self.styleFamily = ""
                self.styleColumnWidth = ""
                self.styleColor = ""
                self.styleBackgroundColor = ""
                self.nameToCode = {}

            def startDocument(self):
                doc.startDocument()

            def endDocument(self):
                doc.endDocument()

            def startElement(self, name, attrs):
                global stackIdx
                stackIdx = stackIdx + 1
                global stackMax
                stackMax = max(stackMax, stackIdx)

                if name == "office:font-face-decls" \
                    or name == "office:master-styles" \
                    or name == "office:styles" \
                    or name == "office:automatic-styles" \
                    or name == "office:forms" \
                    or name == "meta:creation-date" \
                    or name == "meta:print-date" \
                    or name == "meta:editing-cycles" \
                    or name == "meta:editing-duration" \
                    or name == "dc:date" \
                    :
                    #or name == "dc:language" \/
                    #stack.append(False)
                    stack[stackIdx] = False
                    if name == "office:styles":
                        self.isOfficeStyles = True
                    return
                if name == "config:config-item":
                    if attrs["config:name"] in {
                            "VisibleAreaTop", "VisibleAreaLeft", "VisibleAreaWidth", "VisibleAreaHeight",
                            "CursorPositionX", "CursorPositionY",
                            "ActiveSplitRange",
                            #"PositionLeft", "PositionRight",
                            "PositionTop", "PositionBottom",
                            "ZoomType", "ZoomValue", "PageViewZoomValue",
                            "ShowGrid", "AnchoredTextOverflowLegacy",
                            "LegacySingleLineFontwork", "ConnectorUseSnapRect",
                            "IgnoreBreakAfterMultilineField", "ActiveTable",
                            "HorizontalScrollbarWidth", "ShowPageBreakPreview",
                            "ShowZeroValues", "ShowNotes", "ShowFormulasMarks",
                            "ShowGrid", "GridColor", "ShowPageBreaks", "FormulaBarHeight",
                            "HasSheetTabs", "IsOutlineSymbolsSet", "IsValueHighlightingEnabled",
                            "IsSnapToRaster", "RasterIsVisible",
                            "RasterResolutionX", "RasterResolutionY",
                            "RasterSubdivisionX", "RasterSubdivisionY",
                            "IsRasterAxisSynchronized", "AnchoredTextOverflowLegacy",
                            "LegacySingleLineFontwork", "ConnectorUseSnapRect",
                            "IgnoreBreakAfterMultilineField"
                        }:
                        #stack.append(False)
                        stack[stackIdx] = False
                        return
                if name == "config:config-item-set":
                    if attrs["config:name"] == "ooo:configuration-settings":
                        #stack.append(False)
                        stack[stackIdx] = False
                        return

                if name == "office:document":
                    for item in attrs.items():
                        if item[0].startswith("xmlns:"):
                            ET.register_namespace(item[0][6:], item[1])
                # </office:styles>
                #
                # <style:style style:name="co???" style:family="table-column">
                #  <style:table-column-properties style:column-width="1.039in" />
                # </style:style>
                #
                # <style:style style:name="ce???" style:family="table-cell">
                #  <style:table-cell-properties fo:background-color="#ffff00"
                #  <style:text-properties fo:color="#000000"
                # </style:style>
                #
                # </office:automatic-styles>
                elif name == "style:style" and not self.isOfficeStyles:
                    self.styleName = attrs["style:name"]
                    self.styleFamily = attrs["style:family"]

                elif name == "style:table-column-properties" and not self.isOfficeStyles:
                    if "style:column-width" in attrs:
                        self.styleColumnWidth = attrs["style:column-width"]

                elif name == "style:table-cell-properties" and not self.isOfficeStyles:
                    if "fo:background-color" in attrs:
                        self.styleBackgroundColor = attrs["fo:background-color"]
                        if self.styleBackgroundColor == "transparent":
                            self.styleBackgroundColor = ""

                elif name == "style:text-properties" and not self.isOfficeStyles:
                    if "fo:color" in attrs:
                        self.styleColor = attrs["fo:color"]

                elif name == "table:table":
                    tableName = attrs["table:name"]
                    attrs = {}
                    attrs["table:name"] = tableName

                elif name == "table:table-column":
                    if "table:style-name" in attrs \
                            and attrs["table:style-name"] in self.nameToCode:
                        oldName = attrs["table:style-name"]
                        attrs = {}
                        attrs["table:style-name"] = self.nameToCode[oldName]
                    else:
                        attrs = {}

                elif name == "table:table-row":
                    attrs = {}

                elif name == "table:table-cell":
                    numberColumnsRepeated = ""
                    if "table:number-columns-repeated" in attrs:
                        numberColumnsRepeated = attrs["table:number-columns-repeated"]
                    styleName = ""
                    if "table:style-name" in attrs \
                            and attrs["table:style-name"] in self.nameToCode:
                        styleName = self.nameToCode[attrs["table:style-name"]]
                    attrs = {}
                    if len(numberColumnsRepeated):
                        attrs["table:number-columns-repeated"] = numberColumnsRepeated
                    if len(styleName):
                        attrs["table:style-name"] = styleName

                if not stack[stackIdx - 1]:
                    #stack.append(False)
                    stack[stackIdx] = False
                    return
                doc.startElement(name, attrs)

                stack[stackIdx] = True
                global countStart
                countStart = countStart + 1
                #print(f"Start Element: {name}")

            def endElement(self, name):
                if name == "office:styles":
                    self.isOfficeStyles = False
                elif name == "style:style" and len(self.styleName):
                    if len(self.styleColumnWidth) or len(self.styleColor) \
                            or len(self.styleBackgroundColor):
                        code = hashlib.blake2s((self.styleColumnWidth
                                    + ";" + self.styleColor
                                    + ";" + self.styleBackgroundColor
                                    ).encode()).hexdigest()[:16]
                        self.styles[code] = {}
                        self.styles[code]["family"] = self.styleFamily
                        if len(self.styleColumnWidth):
                            self.styles[code]["column-width"] = self.styleColumnWidth
                        if len(self.styleColor):
                            self.styles[code]["color"] = self.styleColor
                        if len(self.styleBackgroundColor):
                            self.styles[code]["background-color"] = self.styleBackgroundColor
                        self.nameToCode[self.styleName] = code
                        print(f"from={self.styleName} to={code} style={self.styles[code]}")
                    self.styleName = ""
                    self.styleFamily = ""
                    self.styleColumnWidth = ""
                    self.styleColor = ""
                    self.styleBackgroundColor = ""
                elif name == "office:automatic-styles":
                    doc.startElement("office:automatic-styles", {})
                    #print(self.styles)
                    for style in sorted(self.styles.items()):
                        doc.startElement("style:style", {
                            "style:name": style[0],
                            "style:family": style[1]["family"]
                        })
                        if "column-width" in style[1]:
                            doc.startElement("style:table-column-properties", {
                                "style:column-width": style[1]["column-width"]
                            })
                            doc.endElement("style:table-column-properties")
                        if "color" in style[1]:
                            doc.startElement("style:text-properties", {
                                "fo:color": style[1]["color"]
                            })
                            doc.endElement("style:text-properties")
                        if "background-color" in style[1]:
                            doc.startElement("style:table-cell-properties", {
                                "fo:background-color": style[1]["background-color"]
                            })
                            doc.endElement("style:table-cell-properties")
                        doc.endElement("style:style")
                    doc.endElement("office:automatic-styles")


                global stackIdx
                if stack[stackIdx]:
                    doc.endElement(name)
                    global countEnd
                    countEnd = countEnd + 1
                    #print(f"End Element: {name}")
                #stack.pop()
                stackIdx = stackIdx - 1

            def characters(self, content):
                global stackIdx
                if stack[stackIdx]:
                    doc.characters(content)
                    global countChars
                    countChars = countChars + 1
                    #content = content.replace("\n", r"\n")
                    #content = content.replace(" ", "_")
                    #print(f"Characters: {len(content)}:{content}")

        handler = MyContentHandler()
        xml.sax.parse(fileName, handler)

        print(f"countStart={countStart} countChars={countChars} countEnd={countEnd} stackMax={stackMax}")

        dom = ET.ElementTree(ET.fromstring(buffer.getvalue()))
        ET.indent(dom, space=" ", level=0)
        #dom.write("_" + fileName, encoding="UTF-8", xml_declaration=True)
        dom.write(fileName, encoding="UTF-8", xml_declaration=True)
