#! /usr/bin/env python
#coding=utf-8

import os
import argparse

# tab对应空格数
TAB_SPACE_NUM = 4

# 字段行
FIELD_NAME_ROW = 0
FIELD_DESC_ROW = 1
FIELD_TYPE_ROW = 2
FIELD_FILTER_ROW = 3
FIELD_ROW_MAX = 4

# 标签
FIELD_NAME_TAG = "[name]"
FIELD_DESC_TAG = "[desc]"
FIELD_TYPE_TAG = "[type]"
FIELD_FILTER_TAG = "[filter]"

# 数组标识以及分隔符
ARRAY_IDENTITY = "|array"
MESSAGE_IDENTITY = "|message"
ARRAY_SEPERATOR = ";"

# 数据起始行
DATA_START_ROW = 4

# 整型数值类型组
INT_TYPE_GROUP = ("int32", "uint32", "int64", "uint64", "sint32", "sint64", "fixed32", "fixed64", "sfixed32", "sfixed64")



"""
日志辅助类
"""
class LogHelper:
    m_logger = None
    m_bFileInited = False

    @staticmethod
    def IsLogEnabled():
        return LogHelper.m_logger is not None

    @staticmethod
    def Create():
        if LogHelper.m_logger is not None :
            return

        import logging
        LogHelper.m_logger = logging.getLogger()
        if not LogHelper.m_bFileInited:
            handler = logging.FileHandler("logger.log", encoding="utf-8")
            formatter = logging.Formatter('[%(asctime)s] %(levelname)-9s: %(message)s', '%m-%d %H:%M:%S')
            handler.setFormatter(formatter)
            LogHelper.m_logger.addHandler(handler)
            LogHelper.m_bFileInited = True

        LogHelper.m_logger.setLevel(logging.NOTSET)

    @staticmethod
    def Close():
        if LogHelper.IsLogEnabled():
            import logging
            logging.shutdown()
            LogHelper.m_logger = None

    @staticmethod
    def Info(message, *args, **kwargs):
        if LogHelper.IsLogEnabled():
            LogHelper.m_logger.info(message, *args, **kwargs)

    @staticmethod
    def Warning(message, *args, **kwargs):
        if LogHelper.IsLogEnabled():
            LogHelper.m_logger.warning(message, *args, **kwargs)

    @staticmethod
    def Error(message, *args, **kwargs):
        if LogHelper.IsLogEnabled():
            LogHelper.m_logger.error(message, *args, **kwargs)



"""
翻译器
"""
class SheetTranslator:
    m_filePath = ""
    m_sheetName = ""
    m_packageName = ""
    m_filter = ""
    m_workBook = None
    m_sheet = None
    m_module = None
    # 当前行和列
    m_curRow = 0
    m_curCol = 0
    # 总行数和列数
    m_rowCount = 0
    m_colCount = 0
    # tab数量
    m_tabNum = 0
    # message字段表，格式：{msgName: {fieldName: {fieldDesc:, fieldType:, fieldNo:, isArray:},}}
    m_dictMsgStruct = None
    # 内容
    m_content = None
    # pb结构
    m_itemArr = None


    def Init(self, filePath, packageName, sheetName, sheetId, filter):
        self.m_filePath = filePath
        self.m_sheetName = sheetName
        self.m_packageName = packageName
        self.m_filter = filter
        self.m_workBook = None
        self.m_sheet = None
        self.m_module = None
        self.m_rowCount = 0
        self.m_colCount = 0
        self.m_tabNum = 0
        self.m_dictMsgStruct = {}
        self.m_content = []
        
        return self.__Load(sheetId)


    def __Load(self, sheetId):
        import xlrd
        try:
            self.m_workBook = xlrd.open_workbook(self.m_filePath)
            if self.m_sheetName != None:
                self.m_sheet = self.m_workBook.sheet_by_name(self.m_sheetName)
            else:
                self.m_sheetName = self.m_workBook.sheet_names()[sheetId]
                self.m_sheet = self.m_workBook.sheet_by_index(sheetId)

        except BaseException as e:
            LogHelper.Info(f"Open file '{self.m_filePath}' FAIL! exception: {repr(e)}")
            return False

        if self.m_sheet == None:
            return False

        # 总行数和列数
        self.m_rowCount = len(self.m_sheet.col_values(0))
        self.m_colCount = len(self.m_sheet.row_values(0))
        return True


    def ParseHead(self):
        dictMsgFieldIdx = {}

        # 遍历表头
        for self.m_curCol in range(self.m_colCount):
            fieldName = self.m_sheet.cell_value(FIELD_NAME_ROW, self.m_curCol)
            if fieldName == "":
                self.m_colCount = self.m_curCol
                break

            fieldDesc = self.m_sheet.cell_value(FIELD_DESC_ROW, self.m_curCol)
            fieldType = self.m_sheet.cell_value(FIELD_TYPE_ROW, self.m_curCol)
            fieldFilter = self.m_sheet.cell_value(FIELD_FILTER_ROW, self.m_curCol)

            if self.m_curCol == 0:
                fieldName = fieldName[len(FIELD_NAME_TAG):]
                fieldDesc = fieldDesc[len(FIELD_DESC_TAG):]
                fieldType = fieldType[len(FIELD_TYPE_TAG):]
                fieldFilter = fieldFilter[len(FIELD_FILTER_TAG):]

            LogHelper.Info(f"Row: {str(self.m_curRow)} Col: {str(self.m_curCol)} fieldName: {fieldName} fieldDesc: {fieldDesc} fieldType: {fieldType} fieldFilter: {fieldFilter}")
            if self.m_filter == None or fieldFilter.find(self.m_filter) != -1:
                self.__ParseOneField(fieldName, fieldDesc, fieldType, dictMsgFieldIdx)


    def __ParseOneField(self, fieldName, fieldDesc, fieldType, dictMsgFieldIdx):
        isArray = fieldType.find(ARRAY_IDENTITY) > 0
        if isArray:
            fieldType = fieldType.split(ARRAY_IDENTITY)[0]


        msgStr = self.m_sheetName
        if fieldName.find('.') > 0:
            fieldNameArr = fieldName.split('.')
            typeNameArr = fieldType.split('.')
            if len(fieldNameArr) != len(typeNameArr):
                LogHelper.Error(f"Syntax error, fieldName: {fieldName} fieldType: {fieldType}")
                return

            fieldName = fieldNameArr[-1]
            fieldType = typeNameArr[-1]

            # 为字段的前部分消息结构填充字段
            for index in range(len(fieldNameArr) - 1):
                tmpFieldName = fieldNameArr[index]
                if msgStr not in dictMsgFieldIdx:
                    dictMsgFieldIdx[msgStr] = 0
                    self.m_dictMsgStruct[msgStr] = {}
                
                msgStruct = self.m_dictMsgStruct[msgStr]
                if tmpFieldName not in msgStruct:
                    dictMsgFieldIdx[msgStr] += 1
                    msgStruct[tmpFieldName] = {"fieldDesc": typeNameArr[index], "fieldType": typeNameArr[index], "fieldNo": str(dictMsgFieldIdx[msgStr]), "isArray": False, "isMsg": True,}
                msgStr += "." + typeNameArr[index]
                
        # 添加字段记录
        if msgStr not in dictMsgFieldIdx:
            dictMsgFieldIdx[msgStr] = 0
            self.m_dictMsgStruct[msgStr] = {}

        dictMsgFieldIdx[msgStr] += 1
        msgStruct = self.m_dictMsgStruct[msgStr]
        if fieldName not in msgStruct:
            msgStruct[fieldName] = {"fieldDesc": fieldDesc, "fieldType": fieldType, "fieldNo": str(dictMsgFieldIdx[msgStr]), "isArray": isArray, "isMsg": False,}


    def GenProtoFile(self):
        self.__ResetContent()
        # 生成pb描述信息
        self.__GenPBFileHeader()
        # 生成Msg
        outputedMsg = {self.m_sheetName: True,}
        self.__GenOneMsg(self.m_sheetName, outputedMsg)

        # 生成表结构数组
        self.__GenPBMsgHeader(self.__GetSheetMsgArrName())
        self.m_content.append(" " * (self.m_tabNum * TAB_SPACE_NUM) + f"repeated {self.m_sheetName} arrItems = 1;\n")
        self.__GenPBMsgTail()

        # 写入文件
        self.__WriteToFile(self.__GetPBFileName())


    def __ResetContent(self):
        self.m_content.clear()
        self.m_tabNum = 0


    def __GenOneMsg(self, msgName, outputedMsg):
        if msgName not in self.m_dictMsgStruct:
            return

        # 生成消息头
        self.__GenPBMsgHeader(msgName)

        # 遍历包含字段
        msgStruct = self.m_dictMsgStruct[msgName]
        for key in msgStruct:
            self.__GenOneField(msgName, key, msgStruct[key], outputedMsg)

        # 生成消息尾
        self.__GenPBMsgTail()

    
    def __GenOneField(self, msgName, fieldName, dictField, outputedMsg):
        # 先生成注释
        self.__GenPBComment(dictField["fieldDesc"])

        # 消息类型，检查是否已生成代码
        fieldType = dictField["fieldType"]
        subMsgName = msgName + "." + fieldType
        if subMsgName in self.m_dictMsgStruct and subMsgName not in outputedMsg:
            self.__GenOneMsg(subMsgName, outputedMsg)
            outputedMsg[fieldType] = True
        
        # 缩进
        spaceNum = self.m_tabNum * TAB_SPACE_NUM
        self.m_content.append(" " * spaceNum)

        # 数组
        if int(dictField["isArray"]) == 1:
            self.m_content.append("repeated ")

        self.m_content.append(f"{fieldType} {fieldName} = {dictField['fieldNo']};\n")


    def __GetPBFileName(self):
        return self.m_sheetName + ".proto"

    def __GetPBBinFileName(self):
        return self.m_sheetName + ".bytes"

    def __GetModuleName(self):
        return self.m_sheetName.lower() + "_pb2"

    def __GetSheetMsgArrName(self):
        return self.m_sheetName + "Array"


    def __GenPBFileHeader(self):
        self.m_content.append("/**\n")
        self.m_content.append(f"* @file:   {self.__GetPBFileName()}\n")
        self.m_content.append("* @author: lfeng \n")
        self.m_content.append("* @brief:  This file is auto generated by xls2proto, DO NOT MODIFY IT!\n")
        self.m_content.append("*/\n\n")
        self.m_content.append(f"syntax=\"proto3\";\npackage {self.m_packageName};\n\n")


    def __GenPBMsgHeader(self, msgName):
        spaceNum = self.m_tabNum * TAB_SPACE_NUM
        msgNameArr = msgName.split('.')
        self.m_content.append(" " * spaceNum + f"message {msgNameArr[len(msgNameArr) - 1]}\n")
        self.m_content.append(" " * spaceNum + "{\n")
        self.m_tabNum += 1


    def __GenPBMsgTail(self):
        self.m_tabNum -= 1
        self.m_content.append(" "*(self.m_tabNum * TAB_SPACE_NUM) + "}\n\n")


    def __GenPBComment(self, comment):
        spaceNum = self.m_tabNum * TAB_SPACE_NUM
        self.m_content.append(" " * spaceNum + "/* ")

        # 仅一个换行符，不换行
        newLineCount = comment.count("\n")
        if newLineCount <= 1:
            self.m_content.append(comment + " */\n")
            return

        # 非换行符结尾，补上换行符
        if comment[-1] != '\n':
            comment += "\n"
            newLineCount += 1
        
        # 换行符替换为换行符+缩进
        comment = comment.replace("\n", "\n" + " " * spaceNum, newLineCount-1)
        self.m_content.append(comment)
        self.m_content.append(" " * spaceNum + "*/\n")

    
    def __WriteToFile(self, fileName, encoding="utf-8"):
        with open(fileName, "w", encoding=encoding) as f:
            f.writelines(self.m_content)

    def __WriteBinaryToFile(self, content, fileName):
        with open(fileName, "wb") as f:
            f.write(content)

    def LoadProtoModule(self):
        moduleName = self.__GetModuleName()
        try:
            import sys
            import os
            
            sys.path.append(os.getcwd())
            os.system(f"protoc --python_out=./ {self.__GetPBFileName().lower()}")
            exec(f"from {moduleName} import *")
            self.m_module = sys.modules[moduleName]
        except BaseException as e:
            print(f"load module '{moduleName}' failed")
            raise e
    

    def ParseData(self, luaContent):
        LogHelper.Info(f"Begin parsing file '{self.m_filePath}'")
        if luaContent != None:
            luaContent.append(f"{self.m_packageName}.{self.__GetSheetMsgArrName()}" + " = {}\n")

        self.m_itemArr = getattr(self.m_module, self.__GetSheetMsgArrName())()
        actualRow = 1
        for self.m_curRow in range(FIELD_ROW_MAX, self.m_rowCount):
            item = self.m_itemArr.arrItems.add()
            if self.__ParseLine(item, luaContent, actualRow):
                actualRow += 1


    def __ParseLine(self, item, luaContent, actualRow):
        dictRowTabInited = {}
        prefix = f"{self.m_packageName}.{self.__GetSheetMsgArrName()}[{actualRow}]"

        if luaContent != None:
            luaContent.append(f"{prefix}" + " = {}\n")

        rowValid = False
        for self.m_curCol in range(0, self.m_colCount):
            if self.m_sheet.cell_type(self.m_curRow, self.m_curCol) == 0:
                continue

            rowValid = True
            filter = self.m_sheet.cell_value(FIELD_FILTER_ROW, self.m_curCol)
            if self.m_filter == None or filter.find(self.m_filter) != -1:
                self.__ParseOneFieldData(item, dictRowTabInited, prefix, luaContent)

        if not rowValid:
            luaContent.pop()
        return rowValid


    def __ParseOneFieldData(self, item, dictRowTabInited, prefix, luaContent):
        fieldValue = self.m_sheet.cell_value(self.m_curRow, self.m_curCol)
        if fieldValue == None:
            return
        
        # 字段名
        fieldName = self.m_sheet.cell_value(FIELD_NAME_ROW, self.m_curCol)
        if self.m_curCol == 0:
            fieldName = fieldName[len(FIELD_NAME_TAG):]
        fieldNameArr = fieldName.split('.')
        fieldName = fieldNameArr[-1]

        # 字段类型
        fieldType = self.m_sheet.cell_value(FIELD_TYPE_ROW, self.m_curCol)
        if self.m_curCol == 0:
            fieldType = fieldType[len(FIELD_TYPE_TAG):]
        fieldTypeArr = fieldType.split('.')

        tmpMsg = item
        msgName = self.m_sheetName
        preFieldName = ""
        for index in range(len(fieldTypeArr) - 1):
            msgName += "." + fieldTypeArr[index]
            preFieldName += "." + fieldNameArr[index]

            if luaContent != None and preFieldName not in dictRowTabInited:
                dictRowTabInited[preFieldName] = True
                luaContent.append(f"{prefix}{preFieldName} = " + "{}\n")

            tmpMsg = tmpMsg.__getattribute__(fieldNameArr[index])
            if not tmpMsg:
                LogHelper.Error(f"Unexpected fieldName: {fieldNameArr[index]} msgName: {msgName}")
                return

        if len(msgName) <= 0 and msgName not in self.m_dictMsgStruct:
            LogHelper.Error(f"Unknown msg name: {msgName}")
            return
        
        msgStruct = self.m_dictMsgStruct[msgName]
        if fieldName not in msgStruct:
            LogHelper.Error(f"Unexpected fieldName: {fieldName} msgName: {msgName}")
            return

        fieldInfo = msgStruct[fieldName]
        fieldType = fieldInfo["fieldType"]
        if fieldInfo["isArray"]:

            if luaContent != None:
                luaContent.append(f"{prefix}{preFieldName}.{fieldName} = " + "{")

            fieldValueArr = fieldValue.split(ARRAY_SEPERATOR)
            for value in fieldValueArr:
                value = self.__GetFieldValue(fieldType, value)
                if value != None:
                    tmpMsg.__getattribute__(fieldName).append(value)
                    self.__AppendLuaFieldValue(luaContent, fieldType, value, ", ")

            if luaContent != None:
                luaContent.append("}\n")
            return

        if fieldInfo["isMsg"]:
            LogHelper.Error(f"Message field can not be set directly: {fieldName} msgName: {msgName}")
            return
        

        fieldValue = self.__GetFieldValue(fieldType, fieldValue)
        if fieldValue != None:
            if fieldType == "bytes":
                tmpMsg.__getattribute__(fieldName).append(fieldValue)
            else:
                tmpMsg.__setattr__(fieldName, fieldValue)
            
            if luaContent:
                luaContent.append(f"{prefix}{preFieldName}.{fieldName} = ")
                self.__AppendLuaFieldValue(luaContent, fieldType, fieldValue, "\n")


    def __AppendLuaFieldValue(self, luaContent, fieldType, fieldValue, suffix):
        if luaContent == None:
            return
        
        if fieldType == "string" or fieldType == "bytes":
            luaContent.append(f"\"{fieldValue}\"{suffix}")
            return

        luaContent.append(f"{fieldValue}{suffix}")

    
    def __GetFieldValue(self, fieldType, fieldValue):
        if len(str(fieldValue)) == 0:
            return None

        try:
            # double 或 float
            if fieldType == "double" or fieldType == "float":
                return float(fieldValue)

            # bytes
            if fieldType == "bytes":
                fieldValue = str(fieldValue).encode("utf-8")
                if len(fieldValue) <= 0:
                    return None
                else:
                    return fieldValue

            # 整型数值
            for intType in INT_TYPE_GROUP:
                if fieldType == intType:
                    return int(fieldValue)

            return fieldValue

        except Exception as e:
            LogHelper.Error(f"__GetFieldValue FAIL! {repr(e)}")
            return None
    

    def GenPBBinFile(self, fileExt):
        data = self.m_itemArr.SerializeToString()
        self.__WriteBinaryToFile(data, f"{self.m_sheetName}.{fileExt}")
        

    def GenLuaFile(self, content, fileExt):
        if content == None:
            return

        self.__ResetContent()
        self.__GenLuaFileHeader(fileExt)
        self.m_content.extend(content)
        self.__GenLuaFileTail()
        self.__WriteToFile(f"{self.m_packageName.lower()}.{fileExt}")


    def __GenLuaFileHeader(self, fileExt):
        self.m_content.append("--[[\n")
        self.m_content.append(f"* @file:   {self.m_packageName.lower()}.{fileExt}\n")
        self.m_content.append("* @author: lfeng \n")
        self.m_content.append("* @brief:  This file is auto generated by xls2proto, DO NOT MODIFY IT!\n")
        self.m_content.append("]]\n\n")
        self.m_content.append(f"local {self.m_packageName} = " + "{}\n")


    def __GenLuaFileTail(self):
        self.m_content.append(f"\nreturn {self.m_packageName}\n\n")



"""
主逻辑
"""
def __OneFileRoutine(translator, luaContent, args, fileName):
    if not translator.Init(fileName, args.package, args.sheet_name, args.sheet_id, args.filter):
        return

    translator.ParseHead()
    translator.GenProtoFile()
    translator.LoadProtoModule()
    translator.ParseData(luaContent)
    translator.GenPBBinFile(args.bin_ext)


def __TraverseFiles(translator, luaContent, args, path):
    files = os.listdir(path)
    for file in files:
        if not os.path.isdir(path + os.sep + file):
            if file.endswith(".xlsx") or file.endswith(".xls"):
                __OneFileRoutine(translator, luaContent, args, path + os.sep + file)
        else:
            __TraverseFiles(translator, luaContent, args, path + os.sep + file)


def __MainRoutine(args):
    # LogHelper Create
    LogHelper.Create()

    translator = SheetTranslator()
    luaContent = None
    if len(args.lua_ext) > 0:
        luaContent = []
    
    if args.file != None:
        __OneFileRoutine(translator, luaContent, args, args.file)
    else:
        __TraverseFiles(translator, luaContent, args, args.in_path)
    
    translator.GenLuaFile(luaContent, args.lua_ext)


if __name__ == '__main__' :
    parser = argparse.ArgumentParser()
    parser.add_argument("--in_path", default=f".{os.sep}", help="Specify the input path, default is the current path.")
    parser.add_argument("--out_path", default=f".{os.sep}", help="Specify the output path, default is the current path.")
    parser.add_argument("--filter", help="Specify the field filter.")
    parser.add_argument("--file", help="Specify one excel file.")
    parser.add_argument("--sheet_name", help="Specify the sheet name to use.")
    parser.add_argument("--sheet_id", help="Specify the sheet id to use, start at 0.", type=int, default=0)
    parser.add_argument("--package", help="Specify the package name.", required=True)
    parser.add_argument("--lua_ext", help="Specify the lua file extension.", default="lua")
    parser.add_argument("--bin_ext", help="Specify the binary file extension.", default="bytes")
    __MainRoutine(parser.parse_args())

