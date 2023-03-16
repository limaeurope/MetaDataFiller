import sys
from lxml import etree as et



# fixKeys.attrib["val"] = "1"
#
# fk1 = et.SubElement(fixKeys, "Fix1")
# uikey = et.SubElement(fk1, "UIKey")
# uikey.text = et.CDATA("Project Name")
#
# dbkey = et.SubElement(fk1, "DBKey")
# dbkey.text = et.CDATA("PROJECTNAME")
#
# value = et.SubElement(fk1, "value")
# value.text = et.CDATA("<####>")


customKeys.attrib["val"] = "0"


class KeyBase:
    def __init__(self, p_sUIKey, p_sDBKey, p_sValue):
        self.sUIKey = p_sUIKey
        self.sDBKey = p_sDBKey
        self.sValue = p_sValue

    def toEtree(self):
        res = et.Element("_")
        uikey = et.SubElement(res, "UIKey")
        uikey.text = et.CDATA(self.sUIKey)

        dbkey = et.SubElement(res, "DBKey")
        dbkey.text = et.CDATA(self.sDBKey)

        value = et.SubElement(res, "value")
        value.text = et.CDATA(self.sValue)

        return res


class FixKey(KeyBase):
    iFixKeys = 0

    def __init__(self, p_sUIKey, p_sDBKey, p_sValue):
        self.iFixKeys += 1
        super(FixKey, self).__init__(p_sUIKey, p_sDBKey, p_sValue)


def createXML(p_iRow="100", p_sXLSPath=""):
    projectInfo = et.Element("ProjectInfo")

    ver = et.SubElement(projectInfo, "Version")
    ver.attrib["val"] = "3"

    fixKeys = []
    customKeys = []

    _fixKeys = et.SubElement(projectInfo, "FixKeys")

    _iFk = 1
    for fk in fixKeys:
        _fk = et.SubElement(_fixKeys, "Fix" + str(_iFk))
        __fk
        _iFk += 1

    # customKeys = et.SubElement(projectInfo, "CustomKeys")



    tree = et.ElementTree(projectInfo)
    et.indent(tree, " ")
    tree.write("output.xml", xml_declaration=True, encoding="UTF-8")

if __name__ == "__main__":
    createXML(sys.argv[1], sys.argv[2])