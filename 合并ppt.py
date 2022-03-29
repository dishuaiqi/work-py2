from win32com.client import Dispatch
def join_ppt():
    ppt = Dispatch('PowerPoint.Application')
    ppt.Visible = 1
    ppt.DisplayAlerts = 0
    #拿到所有的ppt地址！
    alldirs=[]
    pptA = ppt.Presentations.Open(alldirs[0])
    for dir in alldirs[1:len(alldirs):1]:
        pptB = ppt.Presentations.Open(dir)
        numB=len(pptB.Slides)
        for i in range(1,numB+1):
            pptB.Slides(i).Copy()
            pptA.Slides.Paste()
        pptB.Close()
    #设置合并后ppt的路径！
    pptA.SaveAs('合并后ppt的路径')
    pptA.Close()
    ppt.Quit()