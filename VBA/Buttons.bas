Sub Printing()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' شيت فعلي را ذخيره کنيد
    Set ws = ActiveSheet
    
    ' يافتن آخرين رديف غير خالي در ستون B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row - 4
    
    ' تنظيم محدوده چاپ
    ws.PageSetup.PrintArea = "$B$2:$J$" & lastRow
    
    ' تنظيمات چاپ
    With ws.PageSetup
        .Orientation = xlLandscape ' جهت صفحه
        .PaperSize = xlPaperA4 ' اندازه کاغذ
        .PrintTitleRows = "$2:$12" ' رديف‌هايي که در همه صفحات تکرار مي‌شوند
        .FirstPageNumber = xlAutomatic ' شماره صفحه شروع
        .CenterFooter = "Page &P of &N" ' مقدار فوتر (شماره صفحه)
        .LeftMargin = Application.InchesToPoints(0.551181102362205) ' 1.4cm
        .RightMargin = Application.InchesToPoints(0.551181102362205) ' 1.4cm
        .TopMargin = Application.InchesToPoints(0.275590551181102) ' 0.7cm
        .BottomMargin = Application.InchesToPoints(0.551181102362205) ' 1.4cm
        .HeaderMargin = Application.InchesToPoints(0.275590551181102) ' 0.7cm
        .FooterMargin = Application.InchesToPoints(0.275590551181102) ' 0.7cm
        .PrintQuality = 600 ' DPI
        .CenterHorizontally = True ' تراز مرکز
        .FitToPagesWide = 1 ' عرض صفحه
        .FitToPagesTall = False ' تنظيم ارتفاع صفحه را غير فعال کنيد
        .ScaleWithDocHeaderFooter = True ' مقياس‌بندي سرصفحه و پاصفحه
        .AlignMarginsHeaderFooter = True ' تراز کردن حاشيه‌هاي سرصفحه و پاصفحه
    End With
    
    ' چاپ شيت فعلي
    ws.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    
    ' پاک کردن محدوده چاپ
    ws.PageSetup.PrintArea = ""
End Sub
