import adsk.core, adsk.fusion, traceback
import adsk.drawing
import time
from tkinter import Tk
from tkinter.filedialog import askdirectory


_app = adsk.core.Application.cast(None)
_ui = adsk.core.UserInterface.cast(None)
handlers = []

_exportPDFFolder = askdirectory(title="Select Folder")
# _exportPDFFolder = "C:/Users/Vinicius Dias/Desktop"
# _exportPDFFolder = _ui.inputBox(title="Coloque o caminho da pasta: ", prompt="Caminho")


def run(context):
    try:
        global _app, _ui
        _app = adsk.core.Application.get()
        _ui = _app.userInterface

        # get f2d datafile
        datafile = None
        # _ui.messageBox(str(_app.data.activeFolder))
        # for df in _app.data.activeProject.rootFolder.dataFiles:
        for df in _app.data.activeFolder.dataFiles:
            if df.fileExtension == "f2d":
                # _ui.messageBox(str(df.name))
                datafile = df
                create_pdf(datafile)

        # check datafile
        if not datafile:
            _ui.messageBox(
                'Abort because the "f2d" file cannot be found in the rootFolder of activeProject.'
            )
            return
    except:
        if _ui:
            _ui.messageBox("Failed:\n{}".format(traceback.format_exc()))


def getTaskList():
    adsk.doEvents()
    tasks = _app.executeTextCommand("Application.ListIdleTasks").split("\n")
    return [s.strip() for s in tasks[2:-1]]


def create_pdf(datafile):
    try:
        # open doc
        docs = _app.documents
        drawDoc: adsk.drawing.DrawingDocument = docs.open(datafile)

        # Tasks to be checked.
        targetTasks = [
            "DocumentFullyOpenedTask",
            "Nu::AnalyticsTask",
            "CheckValidationTask",
            "InvalidateCommandsTask",
        ]

        # check start task
        if not targetTasks[0] in getTaskList():
            _ui.messageBox("Task not found : {}".format(targetTasks[0]))
            return

        # Check the task and determine if the Document is Open.
        for targetTask in targetTasks:
            while True:
                time.sleep(0.1)
                if not targetTask in getTaskList():
                    break

        # export PDF
        expPDFpath = _exportPDFFolder + "/" + drawDoc.name + ".pdf"

        draw: adsk.drawing.Drawing = drawDoc.drawing
        pdfExpMgr: adsk.drawing.DrawingExportManager = draw.exportManager

        pdfExpOpt: adsk.drawing.DrawingExportOptions = pdfExpMgr.createPDFExportOptions(
            expPDFpath
        )
        pdfExpOpt.openPDF = True
        pdfExpOpt.useLineWeights = True

        pdfExpMgr.execute(pdfExpOpt)

        # close doc
        drawDoc.close(False)
    except:
        if _ui:
            _ui.messageBox("Failed:\n{}".format(traceback.format_exc()))
