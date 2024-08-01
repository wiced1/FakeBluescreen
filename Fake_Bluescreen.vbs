set objIE = CreateObject("InternetExplorer.Application")
objIE.Navigate "https://www.ravbug.com/bsod/bsod10/"
objIE.FullScreen = True
objIE.Visible = True