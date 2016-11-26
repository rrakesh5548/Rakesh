Set QTP = CreateObject("QuickTest.Application")
QTP.Launch
'QTP.Visible = True

QTP.Open "D:\QTP Scripts\FaceBook",True

Set qtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")
qtpResultsOpt.ResultsLocation = "C:\Users\Public\Documents\Output"

QTP.Test.Run qtpResultsOpt

QTP.Test.Close
QTP.Quit
