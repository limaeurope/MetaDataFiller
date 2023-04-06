import os.path

from lxml import etree as et
import xlrd
import re

_A_ =  0;   _B_ =  1;   _C_ =  2;   _D_ =  3;   _E_ =  4
_F_ =  5;   _G_ =  6;   _H_ =  7;   _I_ =  8;   _J_ =  9
_K_ = 10;   _L_ = 11;   _M_ = 12;   _N_ = 13;   _O_ = 14
_P_ = 15;   _Q_ = 16;   _R_ = 17;   _S_ = 18;   _T_ = 19
_U_ = 20;   _V_ = 21;   _W_ = 22;   _X_ = 23;   _Y_ = 24
_Z_ = 25


class KeyBase:
    def __init__(self, p_sUIKey, p_sDBKey, p_sValue, p_sName):
        self.sUIKey = p_sUIKey
        self.sDBKey = p_sDBKey
        self.sValue = p_sValue
        self.sName = p_sName

    def toEtree(self):
        res = et.Element(self.sName)
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
        FixKey.iFixKeys += 1
        super(FixKey, self).__init__(p_sUIKey, p_sDBKey, p_sValue, f"Fix{str(FixKey.iFixKeys)}")


class CustomKey(KeyBase):
    iCustomKeys = 0

    def __init__(self, p_sUIKey, p_sDBKey, p_sValue):
        CustomKey.iCustomKeys += 1
        super(CustomKey, self).__init__(p_sUIKey, p_sDBKey, p_sValue, f"Custom{str(CustomKey.iCustomKeys)}")


class SourceRow:
    def __init__(self, p_iRow, p_sXLSPath):
        self.template = xlrd.open_workbook(p_sXLSPath)
        self.sheet = self.template.sheet_by_index(0)
        self.row = self.sheet.row(p_iRow)

    def __getitem__(self, item):
        if isinstance(val := self.row[item].value, str):
            return val
        elif isinstance(val, float):
            return str(int(val))
        else:
            return str(val)


def createXML(p_iRow="14", p_sXLSPath=r"C:\Users\samu.karli\OneDrive - LIMA Design Kft\HU22-036-TE_Helyszin-Klaszterek.xls"):
    sourceRow = SourceRow(int(p_iRow), p_sXLSPath)

    if sourceRow[_B_] not in {"ÉA", "ÉM"} or not re.match(r'\d+', str(sourceRow[_G_])):
        return

    projectInfo = et.Element("ProjectInfo")

    ver = et.SubElement(projectInfo, "Version")
    ver.attrib["val"] = "3"

    FixKey.iFixKeys = 0
    CustomKey.iCustomKeys = 0

    fixKeys = [FixKey("Project Name", "PROJECTNAME", "<####>"),
               FixKey("Client Company", "CLIENTCOMPANY", ""),
               FixKey("Site Name", "SITE_NAME", sourceRow[_H_]),
               FixKey("Site ID", "SITE_ID", sourceRow[_P_]),
               FixKey("Site Address1", "SITEADDRESS1", sourceRow[_L_]),
               FixKey("Site City", "SITECITY", sourceRow[_I_]),
               FixKey("Site Postcode", "SITEPOSTCODE", sourceRow[_J_]),
               ]

    customKeys = [
                  CustomKey("CAD Drafter Responsible Name", "autotext-PROJECT-3F27A8D0-6FBE-4EF8-A3AA-8C0D4D61DDAE", ""),
                  CustomKey("CAD Drafter - Company Name", "autotext-PROJECT-C179E160-3C5D-41A3-8125-1EEFD66B6BB7", "Lima Design Kft."),
                  CustomKey("OM", "autotext-SITE-D1A8B19C-00D2-4D00-898E-C98BA30E26A8", sourceRow[_Q_]),
                  CustomKey("Survey Company Name", "autotext-PROJECT-03B4AF22-C085-4499-9E7A-5E322E8489FE", "Lima Design Kft."),
                  CustomKey("Survey Responsible Name", "autotext-PROJECT-61034D80-84F0-4A5A-BA73-ADF24B043268", "Dormán Bertalan"),
                  ]

    _fixKeys = et.SubElement(projectInfo, "FixKeys")

    _iFk = 1
    for fk in fixKeys:
        _fixKeys.append(fk.toEtree())
        _iFk += 1

    _fixKeys.attrib["val"] = str(_iFk-1)


    _customKeys = et.SubElement(projectInfo, "CustomKeys")

    _iCk = 1
    for fk in customKeys:
        _customKeys.append(fk.toEtree())
        _iCk += 1

    _customKeys.attrib["val"] = str(_iCk-1)

    tree = et.ElementTree(projectInfo)
    et.indent(tree, " ")
    sOutFile = os.path.join("output", sourceRow[_B_] + "-" + str(int(sourceRow[_G_])) + ".xml")

    tree.write(sOutFile, xml_declaration=True, encoding="UTF-8")

if __name__ == "__main__":
    for i in range(6, 274):
        createXML(str(i))

