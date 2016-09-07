wxAutomationObject excelObject,workBooks, workbook, worksheets, sheet, excRange/*, font*/;
if ( !excelObject.GetInstance(wxT("Excel.Application")) )
{
wxLogError(wxT("Could not create Excel object."));
return;
}


// Ensure that Excel is visible
if (!excelObject.PutProperty(wxT("Visible"), true))
{
wxLogError(wxT("Could not make Excel object visible"));
} 


bool ret = excelObject.GetObject(workBooks, wxT("Workbooks"));


if (!ret)
{
return;
}






if (!bFileExist)
{
bSuc = CreateDirectory(strFilePath, NULL);


if (!bSuc)
{
return;
}
}


if (!bFileExist)
{
workBooks.CallMethod(wxT("Add"));
}
else
{
  wxbook = workBooks.CallMethod(wxT("Open"), wxVariant(strFilePath));


  if (wxbook.IsNull())
  {
  workBooks.CallMethod(wxT("Add"));
  }
}


ret = excelObject.GetObject(workbook, wxT("ActiveWorkbook"));
ret = workbook.GetObject(worksheets, wxT("Worksheets"));
ret = worksheets.GetObject(sheet, wxT("Item"), 1, &wxVariant(1));


ret = sheet.GetObject(excRange, wxT("Range"), 1, &wxVariant("A1"));


if (!ret)
{
return;
}

if (bFileExist)
{
excRange.CallMethod(wxT("Clear"));
}


wxString strValue, strMer;
static std::string strColName[] = { "序号" , "编码" , "名称" , "错误类型" };
int nCenter = 
0xFFFFEFF4;
int nLeft = 0xFFFFEFDD;


for (int i = 0; i < 4; ++i)
{
strMer = wxString::Format(wxT("%c%d"), 'A' + i, 1);
ret = sheet.GetObject(excRange, wxT("Range"), 1, &wxVariant(strMer));


if (!ret)
{
return;
}


excRange.PutProperty(wxT("Value"), wxVariant(strColName[i])) ;


switch (i)
{
case 0:
excRange.PutProperty(wxT("HorizontalAlignment"), wxVariant(nCenter)) ;
excRange.PutProperty(wxT("VerticalAlignment"), wxVariant(nCenter)) ;
excRange.PutProperty(wxT("ColumnWidth"), wxVariant(10)) ;
break;
case 1:
excRange.PutProperty(wxT("HorizontalAlignment"), wxVariant(nLeft)) ;  //xlLeft
0xFFFFEFDD
excRange.PutProperty(wxT("VerticalAlignment"), wxVariant(nCenter)) ;
//xlCenter 0xFFFFEFF4
excRange.PutProperty(wxT("ColumnWidth"), wxVariant(20)) ;
break;
case 2:
excRange.PutProperty(wxT("HorizontalAlignment"), wxVariant(nLeft)) ;
excRange.PutProperty(wxT("VerticalAlignment"), wxVariant(nCenter)) ;
excRange.PutProperty(wxT("ColumnWidth"), wxVariant(30)) ;
break;
case 3:
excRange.PutProperty(wxT("HorizontalAlignment"), wxVariant(nLeft)) ;
excRange.PutProperty(wxT("VerticalAlignment"), wxVariant(nCenter)) ;
excRange.PutProperty(wxT("ColumnWidth"), wxVariant(40)) ;
break;
}
}


for (int i = 0; i < nRows; ++i)
{
for (int j = 0; j < 4; ++j)
{
strMer = wxString::Format(wxT("%c%d"), 'A' + j, i + 2);
ret = sheet.GetObject(excRange, wxT("Range"), 1, &wxVariant(strMer));


if (!ret)
{
return;
}

strValue = tbl->GetValue(i, j);
// wxColour fontcolor =tbl->GetAttr(i, j, wxGridCellAttr::Row)->GetTextColour();
// 
// ret = excRange.GetObject(font, wxT("Font"));


if (!ret)
{
return;
}


// font.PutProperty(wxT("Color"), wxVariant((int)fontcolor.GetRGB())) ;


excRange.PutProperty(wxT("NumberFormatLocal"), "@") ;
excRange.PutProperty(wxT("Value"), wxVariant(strValue)) ;
excRange.PutProperty(wxT("ShrinkToFit"), true) ;


switch (j)
{
case 0:
excRange.PutProperty(wxT("HorizontalAlignment"),wxVariant(nCenter)) ;
excRange.PutProperty(wxT("VerticalAlignment"), wxVariant(nCenter)) ;
break;
case 1:
excRange.PutProperty(wxT("HorizontalAlignment"), wxVariant(nLeft)) ;
excRange.PutProperty(wxT("VerticalAlignment"), wxVariant(nCenter)) ;
break;
case 2:
excRange.PutProperty(wxT("HorizontalAlignment"), wxVariant(nLeft)) ;
excRange.PutProperty(wxT("VerticalAlignment"), wxVariant(nCenter)) ;
break;
case 3:
excRange.PutProperty(wxT("HorizontalAlignment"), wxVariant(nLeft)) ;
excRange.PutProperty(wxT("VerticalAlignment"), wxVariant(nCenter)) ;
break;
}

}
}


if (bFileExist)
{
workbook.CallMethod(wxT("Save"));
}
else
{
workbook.CallMethod(wxT("SaveAs"), wxVariant(strFilePath), wxVariant(int(18)));
// xlAddIn = 18
}
