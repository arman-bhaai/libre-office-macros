REM  *****  BASIC  *****



sub FlatGiven
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem //array declaration with 10 variable for keys and values

dim keys(9) as string
dim values(9) as string
rem ----------------------------------------------------------------------
rem //assigned values to the 'Keys' array

keys(0) = "<>"
keys(1) = "<>"
keys(2) = "<>"
keys(3) = "<>"
keys(4) = "<>"
keys(5) = "<>"
keys(6) = "<>"
keys(7) = "<>"
keys(8) = "<>"
keys(9) = "<>"
rem ----------------------------------------------------------------------
rem //assigned values to the 'Values' array

values(0) = ""
values(1) = ""
values(2) = ""
values(3) = ""
values(4) = ""
values(5) = ""
values(6) = ""
values(7) = ""
values(8) = ""
values(9) = ""
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:GoToStartOfDoc", "", 0, Array())

rem ----------------------------------------------------------------------
dim args2(1) as new com.sun.star.beans.PropertyValue
args2(0).Name = "Count"
args2(0).Value = 1
args2(1).Name = "Select"
args2(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args2())

rem ----------------------------------------------------------------------
dim args3(1) as new com.sun.star.beans.PropertyValue
args3(0).Name = "Count"
args3(0).Value = 1
args3(1).Name = "Select"
args3(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args3())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim args5(21) as new com.sun.star.beans.PropertyValue
args5(0).Name = "SearchItem.StyleFamily"
args5(0).Value = 2
args5(1).Name = "SearchItem.CellType"
args5(1).Value = 0
args5(2).Name = "SearchItem.RowDirection"
args5(2).Value = true
args5(3).Name = "SearchItem.AllTables"
args5(3).Value = false
args5(4).Name = "SearchItem.SearchFiltered"
args5(4).Value = false
args5(5).Name = "SearchItem.Backward"
args5(5).Value = false
args5(6).Name = "SearchItem.Pattern"
args5(6).Value = false
args5(7).Name = "SearchItem.Content"
args5(7).Value = false
args5(8).Name = "SearchItem.AsianOptions"
args5(8).Value = false
args5(9).Name = "SearchItem.AlgorithmType"
args5(9).Value = 0
args5(10).Name = "SearchItem.SearchFlags"
args5(10).Value = 65536
args5(11).Name = "SearchItem.SearchString"
args5(11).Value = keys(0)
args5(12).Name = "SearchItem.ReplaceString"
args5(12).Value = values(0)
args5(13).Name = "SearchItem.Locale"
args5(13).Value = 255
args5(14).Name = "SearchItem.ChangedChars"
args5(14).Value = 2
args5(15).Name = "SearchItem.DeletedChars"
args5(15).Value = 2
args5(16).Name = "SearchItem.InsertedChars"
args5(16).Value = 2
args5(17).Name = "SearchItem.TransliterateFlags"
args5(17).Value = 1280
args5(18).Name = "SearchItem.Command"
args5(18).Value = 3
args5(19).Name = "SearchItem.SearchFormatted"
args5(19).Value = false
args5(20).Name = "SearchItem.AlgorithmType2"
args5(20).Value = 1
args5(21).Name = "Quiet"
args5(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args5())

rem ----------------------------------------------------------------------
dim args6(1) as new com.sun.star.beans.PropertyValue
args6(0).Name = "Count"
args6(0).Value = 1
args6(1).Name = "Select"
args6(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args6())

rem ----------------------------------------------------------------------
dim args7(1) as new com.sun.star.beans.PropertyValue
args7(0).Name = "Count"
args7(0).Value = 1
args7(1).Name = "Select"
args7(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args7())

rem ----------------------------------------------------------------------
dim args8(1) as new com.sun.star.beans.PropertyValue
args8(0).Name = "Count"
args8(0).Value = 1
args8(1).Name = "Select"
args8(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args8())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim args10(21) as new com.sun.star.beans.PropertyValue
args10(0).Name = "SearchItem.StyleFamily"
args10(0).Value = 2
args10(1).Name = "SearchItem.CellType"
args10(1).Value = 0
args10(2).Name = "SearchItem.RowDirection"
args10(2).Value = true
args10(3).Name = "SearchItem.AllTables"
args10(3).Value = false
args10(4).Name = "SearchItem.SearchFiltered"
args10(4).Value = false
args10(5).Name = "SearchItem.Backward"
args10(5).Value = false
args10(6).Name = "SearchItem.Pattern"
args10(6).Value = false
args10(7).Name = "SearchItem.Content"
args10(7).Value = false
args10(8).Name = "SearchItem.AsianOptions"
args10(8).Value = false
args10(9).Name = "SearchItem.AlgorithmType"
args10(9).Value = 0
args10(10).Name = "SearchItem.SearchFlags"
args10(10).Value = 65536
args10(11).Name = "SearchItem.SearchString"
args10(11).Value = keys(1)
args10(12).Name = "SearchItem.ReplaceString"
args10(12).Value = values(1)
args10(13).Name = "SearchItem.Locale"
args10(13).Value = 255
args10(14).Name = "SearchItem.ChangedChars"
args10(14).Value = 2
args10(15).Name = "SearchItem.DeletedChars"
args10(15).Value = 2
args10(16).Name = "SearchItem.InsertedChars"
args10(16).Value = 2
args10(17).Name = "SearchItem.TransliterateFlags"
args10(17).Value = 1280
args10(18).Name = "SearchItem.Command"
args10(18).Value = 3
args10(19).Name = "SearchItem.SearchFormatted"
args10(19).Value = false
args10(20).Name = "SearchItem.AlgorithmType2"
args10(20).Value = 1
args10(21).Name = "Quiet"
args10(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args10())

rem ----------------------------------------------------------------------
dim args11(1) as new com.sun.star.beans.PropertyValue
args11(0).Name = "Count"
args11(0).Value = 1
args11(1).Name = "Select"
args11(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args11())

rem ----------------------------------------------------------------------
dim args12(1) as new com.sun.star.beans.PropertyValue
args12(0).Name = "Count"
args12(0).Value = 1
args12(1).Name = "Select"
args12(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args12())

rem ----------------------------------------------------------------------
dim args13(1) as new com.sun.star.beans.PropertyValue
args13(0).Name = "Count"
args13(0).Value = 1
args13(1).Name = "Select"
args13(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args13())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim args15(21) as new com.sun.star.beans.PropertyValue
args15(0).Name = "SearchItem.StyleFamily"
args15(0).Value = 2
args15(1).Name = "SearchItem.CellType"
args15(1).Value = 0
args15(2).Name = "SearchItem.RowDirection"
args15(2).Value = true
args15(3).Name = "SearchItem.AllTables"
args15(3).Value = false
args15(4).Name = "SearchItem.SearchFiltered"
args15(4).Value = false
args15(5).Name = "SearchItem.Backward"
args15(5).Value = false
args15(6).Name = "SearchItem.Pattern"
args15(6).Value = false
args15(7).Name = "SearchItem.Content"
args15(7).Value = false
args15(8).Name = "SearchItem.AsianOptions"
args15(8).Value = false
args15(9).Name = "SearchItem.AlgorithmType"
args15(9).Value = 0
args15(10).Name = "SearchItem.SearchFlags"
args15(10).Value = 65536
args15(11).Name = "SearchItem.SearchString"
args15(11).Value = keys(2)
args15(12).Name = "SearchItem.ReplaceString"
args15(12).Value = values(2)
args15(13).Name = "SearchItem.Locale"
args15(13).Value = 255
args15(14).Name = "SearchItem.ChangedChars"
args15(14).Value = 2
args15(15).Name = "SearchItem.DeletedChars"
args15(15).Value = 2
args15(16).Name = "SearchItem.InsertedChars"
args15(16).Value = 2
args15(17).Name = "SearchItem.TransliterateFlags"
args15(17).Value = 1280
args15(18).Name = "SearchItem.Command"
args15(18).Value = 3
args15(19).Name = "SearchItem.SearchFormatted"
args15(19).Value = false
args15(20).Name = "SearchItem.AlgorithmType2"
args15(20).Value = 1
args15(21).Name = "Quiet"
args15(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args15())

rem ----------------------------------------------------------------------
dim args16(1) as new com.sun.star.beans.PropertyValue
args16(0).Name = "Count"
args16(0).Value = 1
args16(1).Name = "Select"
args16(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args16())

rem ----------------------------------------------------------------------
dim args17(1) as new com.sun.star.beans.PropertyValue
args17(0).Name = "Count"
args17(0).Value = 1
args17(1).Name = "Select"
args17(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args17())

rem ----------------------------------------------------------------------
dim args18(1) as new com.sun.star.beans.PropertyValue
args18(0).Name = "Count"
args18(0).Value = 1
args18(1).Name = "Select"
args18(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args18())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim args20(21) as new com.sun.star.beans.PropertyValue
args20(0).Name = "SearchItem.StyleFamily"
args20(0).Value = 2
args20(1).Name = "SearchItem.CellType"
args20(1).Value = 0
args20(2).Name = "SearchItem.RowDirection"
args20(2).Value = true
args20(3).Name = "SearchItem.AllTables"
args20(3).Value = false
args20(4).Name = "SearchItem.SearchFiltered"
args20(4).Value = false
args20(5).Name = "SearchItem.Backward"
args20(5).Value = false
args20(6).Name = "SearchItem.Pattern"
args20(6).Value = false
args20(7).Name = "SearchItem.Content"
args20(7).Value = false
args20(8).Name = "SearchItem.AsianOptions"
args20(8).Value = false
args20(9).Name = "SearchItem.AlgorithmType"
args20(9).Value = 0
args20(10).Name = "SearchItem.SearchFlags"
args20(10).Value = 65536
args20(11).Name = "SearchItem.SearchString"
args20(11).Value = keys(3)
args20(12).Name = "SearchItem.ReplaceString"
args20(12).Value = values(3)
args20(13).Name = "SearchItem.Locale"
args20(13).Value = 255
args20(14).Name = "SearchItem.ChangedChars"
args20(14).Value = 2
args20(15).Name = "SearchItem.DeletedChars"
args20(15).Value = 2
args20(16).Name = "SearchItem.InsertedChars"
args20(16).Value = 2
args20(17).Name = "SearchItem.TransliterateFlags"
args20(17).Value = 1280
args20(18).Name = "SearchItem.Command"
args20(18).Value = 3
args20(19).Name = "SearchItem.SearchFormatted"
args20(19).Value = false
args20(20).Name = "SearchItem.AlgorithmType2"
args20(20).Value = 1
args20(21).Name = "Quiet"
args20(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args20())

rem ----------------------------------------------------------------------
dim args21(1) as new com.sun.star.beans.PropertyValue
args21(0).Name = "Count"
args21(0).Value = 1
args21(1).Name = "Select"
args21(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args21())

rem ----------------------------------------------------------------------
dim args22(1) as new com.sun.star.beans.PropertyValue
args22(0).Name = "Count"
args22(0).Value = 1
args22(1).Name = "Select"
args22(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args22())

rem ----------------------------------------------------------------------
dim args23(1) as new com.sun.star.beans.PropertyValue
args23(0).Name = "Count"
args23(0).Value = 1
args23(1).Name = "Select"
args23(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args23())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim args25(21) as new com.sun.star.beans.PropertyValue
args25(0).Name = "SearchItem.StyleFamily"
args25(0).Value = 2
args25(1).Name = "SearchItem.CellType"
args25(1).Value = 0
args25(2).Name = "SearchItem.RowDirection"
args25(2).Value = true
args25(3).Name = "SearchItem.AllTables"
args25(3).Value = false
args25(4).Name = "SearchItem.SearchFiltered"
args25(4).Value = false
args25(5).Name = "SearchItem.Backward"
args25(5).Value = false
args25(6).Name = "SearchItem.Pattern"
args25(6).Value = false
args25(7).Name = "SearchItem.Content"
args25(7).Value = false
args25(8).Name = "SearchItem.AsianOptions"
args25(8).Value = false
args25(9).Name = "SearchItem.AlgorithmType"
args25(9).Value = 0
args25(10).Name = "SearchItem.SearchFlags"
args25(10).Value = 65536
args25(11).Name = "SearchItem.SearchString"
args25(11).Value = keys(4)
args25(12).Name = "SearchItem.ReplaceString"
args25(12).Value = values(4)
args25(13).Name = "SearchItem.Locale"
args25(13).Value = 255
args25(14).Name = "SearchItem.ChangedChars"
args25(14).Value = 2
args25(15).Name = "SearchItem.DeletedChars"
args25(15).Value = 2
args25(16).Name = "SearchItem.InsertedChars"
args25(16).Value = 2
args25(17).Name = "SearchItem.TransliterateFlags"
args25(17).Value = 1280
args25(18).Name = "SearchItem.Command"
args25(18).Value = 3
args25(19).Name = "SearchItem.SearchFormatted"
args25(19).Value = false
args25(20).Name = "SearchItem.AlgorithmType2"
args25(20).Value = 1
args25(21).Name = "Quiet"
args25(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args25())

rem ----------------------------------------------------------------------
dim args26(1) as new com.sun.star.beans.PropertyValue
args26(0).Name = "Count"
args26(0).Value = 1
args26(1).Name = "Select"
args26(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args26())

rem ----------------------------------------------------------------------
dim args27(1) as new com.sun.star.beans.PropertyValue
args27(0).Name = "Count"
args27(0).Value = 1
args27(1).Name = "Select"
args27(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args27())

rem ----------------------------------------------------------------------
dim args28(1) as new com.sun.star.beans.PropertyValue
args28(0).Name = "Count"
args28(0).Value = 1
args28(1).Name = "Select"
args28(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args28())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim args30(21) as new com.sun.star.beans.PropertyValue
args30(0).Name = "SearchItem.StyleFamily"
args30(0).Value = 2
args30(1).Name = "SearchItem.CellType"
args30(1).Value = 0
args30(2).Name = "SearchItem.RowDirection"
args30(2).Value = true
args30(3).Name = "SearchItem.AllTables"
args30(3).Value = false
args30(4).Name = "SearchItem.SearchFiltered"
args30(4).Value = false
args30(5).Name = "SearchItem.Backward"
args30(5).Value = false
args30(6).Name = "SearchItem.Pattern"
args30(6).Value = false
args30(7).Name = "SearchItem.Content"
args30(7).Value = false
args30(8).Name = "SearchItem.AsianOptions"
args30(8).Value = false
args30(9).Name = "SearchItem.AlgorithmType"
args30(9).Value = 0
args30(10).Name = "SearchItem.SearchFlags"
args30(10).Value = 65536
args30(11).Name = "SearchItem.SearchString"
args30(11).Value = keys(5)
args30(12).Name = "SearchItem.ReplaceString"
args30(12).Value = values(5)
args30(13).Name = "SearchItem.Locale"
args30(13).Value = 255
args30(14).Name = "SearchItem.ChangedChars"
args30(14).Value = 2
args30(15).Name = "SearchItem.DeletedChars"
args30(15).Value = 2
args30(16).Name = "SearchItem.InsertedChars"
args30(16).Value = 2
args30(17).Name = "SearchItem.TransliterateFlags"
args30(17).Value = 1280
args30(18).Name = "SearchItem.Command"
args30(18).Value = 3
args30(19).Name = "SearchItem.SearchFormatted"
args30(19).Value = false
args30(20).Name = "SearchItem.AlgorithmType2"
args30(20).Value = 1
args30(21).Name = "Quiet"
args30(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args30())

rem ----------------------------------------------------------------------
dim args31(1) as new com.sun.star.beans.PropertyValue
args31(0).Name = "Count"
args31(0).Value = 1
args31(1).Name = "Select"
args31(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args31())

rem ----------------------------------------------------------------------
dim args32(1) as new com.sun.star.beans.PropertyValue
args32(0).Name = "Count"
args32(0).Value = 1
args32(1).Name = "Select"
args32(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args32())

rem ----------------------------------------------------------------------
dim args33(1) as new com.sun.star.beans.PropertyValue
args33(0).Name = "Count"
args33(0).Value = 1
args33(1).Name = "Select"
args33(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args33())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim args35(21) as new com.sun.star.beans.PropertyValue
args35(0).Name = "SearchItem.StyleFamily"
args35(0).Value = 2
args35(1).Name = "SearchItem.CellType"
args35(1).Value = 0
args35(2).Name = "SearchItem.RowDirection"
args35(2).Value = true
args35(3).Name = "SearchItem.AllTables"
args35(3).Value = false
args35(4).Name = "SearchItem.SearchFiltered"
args35(4).Value = false
args35(5).Name = "SearchItem.Backward"
args35(5).Value = false
args35(6).Name = "SearchItem.Pattern"
args35(6).Value = false
args35(7).Name = "SearchItem.Content"
args35(7).Value = false
args35(8).Name = "SearchItem.AsianOptions"
args35(8).Value = false
args35(9).Name = "SearchItem.AlgorithmType"
args35(9).Value = 0
args35(10).Name = "SearchItem.SearchFlags"
args35(10).Value = 65536
args35(11).Name = "SearchItem.SearchString"
args35(11).Value = keys(6)
args35(12).Name = "SearchItem.ReplaceString"
args35(12).Value = values(6)
args35(13).Name = "SearchItem.Locale"
args35(13).Value = 255
args35(14).Name = "SearchItem.ChangedChars"
args35(14).Value = 2
args35(15).Name = "SearchItem.DeletedChars"
args35(15).Value = 2
args35(16).Name = "SearchItem.InsertedChars"
args35(16).Value = 2
args35(17).Name = "SearchItem.TransliterateFlags"
args35(17).Value = 1280
args35(18).Name = "SearchItem.Command"
args35(18).Value = 3
args35(19).Name = "SearchItem.SearchFormatted"
args35(19).Value = false
args35(20).Name = "SearchItem.AlgorithmType2"
args35(20).Value = 1
args35(21).Name = "Quiet"
args35(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args35())

rem ----------------------------------------------------------------------
dim args36(1) as new com.sun.star.beans.PropertyValue
args36(0).Name = "Count"
args36(0).Value = 1
args36(1).Name = "Select"
args36(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args36())

rem ----------------------------------------------------------------------
dim args37(1) as new com.sun.star.beans.PropertyValue
args37(0).Name = "Count"
args37(0).Value = 1
args37(1).Name = "Select"
args37(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args37())

rem ----------------------------------------------------------------------
dim args38(1) as new com.sun.star.beans.PropertyValue
args38(0).Name = "Count"
args38(0).Value = 1
args38(1).Name = "Select"
args38(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args38())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim args40(21) as new com.sun.star.beans.PropertyValue
args40(0).Name = "SearchItem.StyleFamily"
args40(0).Value = 2
args40(1).Name = "SearchItem.CellType"
args40(1).Value = 0
args40(2).Name = "SearchItem.RowDirection"
args40(2).Value = true
args40(3).Name = "SearchItem.AllTables"
args40(3).Value = false
args40(4).Name = "SearchItem.SearchFiltered"
args40(4).Value = false
args40(5).Name = "SearchItem.Backward"
args40(5).Value = false
args40(6).Name = "SearchItem.Pattern"
args40(6).Value = false
args40(7).Name = "SearchItem.Content"
args40(7).Value = false
args40(8).Name = "SearchItem.AsianOptions"
args40(8).Value = false
args40(9).Name = "SearchItem.AlgorithmType"
args40(9).Value = 0
args40(10).Name = "SearchItem.SearchFlags"
args40(10).Value = 65536
args40(11).Name = "SearchItem.SearchString"
args40(11).Value = keys(7)
args40(12).Name = "SearchItem.ReplaceString"
args40(12).Value = values(7)
args40(13).Name = "SearchItem.Locale"
args40(13).Value = 255
args40(14).Name = "SearchItem.ChangedChars"
args40(14).Value = 2
args40(15).Name = "SearchItem.DeletedChars"
args40(15).Value = 2
args40(16).Name = "SearchItem.InsertedChars"
args40(16).Value = 2
args40(17).Name = "SearchItem.TransliterateFlags"
args40(17).Value = 1280
args40(18).Name = "SearchItem.Command"
args40(18).Value = 3
args40(19).Name = "SearchItem.SearchFormatted"
args40(19).Value = false
args40(20).Name = "SearchItem.AlgorithmType2"
args40(20).Value = 1
args40(21).Name = "Quiet"
args40(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args40())

rem ----------------------------------------------------------------------
dim args41(1) as new com.sun.star.beans.PropertyValue
args41(0).Name = "Count"
args41(0).Value = 1
args41(1).Name = "Select"
args41(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args41())

rem ----------------------------------------------------------------------
dim args42(1) as new com.sun.star.beans.PropertyValue
args42(0).Name = "Count"
args42(0).Value = 1
args42(1).Name = "Select"
args42(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args42())

rem ----------------------------------------------------------------------
dim args43(1) as new com.sun.star.beans.PropertyValue
args43(0).Name = "Count"
args43(0).Value = 1
args43(1).Name = "Select"
args43(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args43())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim args45(21) as new com.sun.star.beans.PropertyValue
args45(0).Name = "SearchItem.StyleFamily"
args45(0).Value = 2
args45(1).Name = "SearchItem.CellType"
args45(1).Value = 0
args45(2).Name = "SearchItem.RowDirection"
args45(2).Value = true
args45(3).Name = "SearchItem.AllTables"
args45(3).Value = false
args45(4).Name = "SearchItem.SearchFiltered"
args45(4).Value = false
args45(5).Name = "SearchItem.Backward"
args45(5).Value = false
args45(6).Name = "SearchItem.Pattern"
args45(6).Value = false
args45(7).Name = "SearchItem.Content"
args45(7).Value = false
args45(8).Name = "SearchItem.AsianOptions"
args45(8).Value = false
args45(9).Name = "SearchItem.AlgorithmType"
args45(9).Value = 0
args45(10).Name = "SearchItem.SearchFlags"
args45(10).Value = 65536
args45(11).Name = "SearchItem.SearchString"
args45(11).Value = keys(8)
args45(12).Name = "SearchItem.ReplaceString"
args45(12).Value = values(8)
args45(13).Name = "SearchItem.Locale"
args45(13).Value = 255
args45(14).Name = "SearchItem.ChangedChars"
args45(14).Value = 2
args45(15).Name = "SearchItem.DeletedChars"
args45(15).Value = 2
args45(16).Name = "SearchItem.InsertedChars"
args45(16).Value = 2
args45(17).Name = "SearchItem.TransliterateFlags"
args45(17).Value = 1280
args45(18).Name = "SearchItem.Command"
args45(18).Value = 3
args45(19).Name = "SearchItem.SearchFormatted"
args45(19).Value = false
args45(20).Name = "SearchItem.AlgorithmType2"
args45(20).Value = 1
args45(21).Name = "Quiet"
args45(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args45())

rem ----------------------------------------------------------------------
dim args46(1) as new com.sun.star.beans.PropertyValue
args46(0).Name = "Count"
args46(0).Value = 1
args46(1).Name = "Select"
args46(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args46())

rem ----------------------------------------------------------------------
dim args47(1) as new com.sun.star.beans.PropertyValue
args47(0).Name = "Count"
args47(0).Value = 1
args47(1).Name = "Select"
args47(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args47())

rem ----------------------------------------------------------------------
dim args48(1) as new com.sun.star.beans.PropertyValue
args48(0).Name = "Count"
args48(0).Value = 1
args48(1).Name = "Select"
args48(1).Value = true

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args48())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())

rem ----------------------------------------------------------------------
dim args50(21) as new com.sun.star.beans.PropertyValue
args50(0).Name = "SearchItem.StyleFamily"
args50(0).Value = 2
args50(1).Name = "SearchItem.CellType"
args50(1).Value = 0
args50(2).Name = "SearchItem.RowDirection"
args50(2).Value = true
args50(3).Name = "SearchItem.AllTables"
args50(3).Value = false
args50(4).Name = "SearchItem.SearchFiltered"
args50(4).Value = false
args50(5).Name = "SearchItem.Backward"
args50(5).Value = false
args50(6).Name = "SearchItem.Pattern"
args50(6).Value = false
args50(7).Name = "SearchItem.Content"
args50(7).Value = false
args50(8).Name = "SearchItem.AsianOptions"
args50(8).Value = false
args50(9).Name = "SearchItem.AlgorithmType"
args50(9).Value = 0
args50(10).Name = "SearchItem.SearchFlags"
args50(10).Value = 65536
args50(11).Name = "SearchItem.SearchString"
args50(11).Value = keys(9)
args50(12).Name = "SearchItem.ReplaceString"
args50(12).Value = keys(9)
args50(13).Name = "SearchItem.Locale"
args50(13).Value = 255
args50(14).Name = "SearchItem.ChangedChars"
args50(14).Value = 2
args50(15).Name = "SearchItem.DeletedChars"
args50(15).Value = 2
args50(16).Name = "SearchItem.InsertedChars"
args50(16).Value = 2
args50(17).Name = "SearchItem.TransliterateFlags"
args50(17).Value = 1280
args50(18).Name = "SearchItem.Command"
args50(18).Value = 3
args50(19).Name = "SearchItem.SearchFormatted"
args50(19).Value = false
args50(20).Name = "SearchItem.AlgorithmType2"
args50(20).Value = 1
args50(21).Name = "Quiet"
args50(21).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args50())


end sub