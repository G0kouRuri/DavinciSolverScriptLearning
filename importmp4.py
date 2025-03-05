import sys
import os


def load_source(module_name, file_path):
    if sys.version_info[0] >= 3 and sys.version_info[1] >= 5:
        import importlib.util

        module = None
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        if spec:
            module = importlib.util.module_from_spec(spec)
        if module:
            sys.modules[module_name] = module
            spec.loader.exec_module(module)
        return module
    else:
        import imp
        return imp.load_source(module_name, file_path)


def GetResolve():
    try:
        # The PYTHONPATH needs to be set correctly for this import statement to work.
        # An alternative is to import the DaVinciResolveScript by specifying absolute path (see ExceptionHandler logic)
        import DaVinciResolveScript as bmd
    except ImportError:
        if sys.platform.startswith("darwin"):
            expectedPath = "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules/"
        elif sys.platform.startswith("win") or sys.platform.startswith("cygwin"):
            import os
            expectedPath = os.getenv('PROGRAMDATA') + "\\Blackmagic Design\\DaVinci Resolve\\Support\\Developer\\Scripting\\Modules\\"
        elif sys.platform.startswith("linux"):
            expectedPath = "/opt/resolve/Developer/Scripting/Modules/"

        # check if the default path has it...
        print("Unable to find module DaVinciResolveScript from $PYTHONPATH - trying default locations")
        try:
            load_source('DaVinciResolveScript', expectedPath + "DaVinciResolveScript.py")
            import DaVinciResolveScript as bmd
            print("import bmd successful!")
        except Exception as ex:
            # No fallbacks ... report error:
            print("Unable to find module DaVinciResolveScript - please ensure that the module DaVinciResolveScript is discoverable by python")
            print("For a default DaVinci Resolve installation, the module is expected to be located in: " + expectedPath)
            print(ex)
            sys.exit()

    return bmd.scriptapp("Resolve")


def TimecodeToFrame(timecode,fps):
    hours,minutes,seconds,frames=map(int,timecode.split(':'))
    return hours*3600*fps+minutes*60*fps+seconds*fps+frames

# 导入文件目录
toAddFilePath="E:/czyapp/Davinci/Projects/Videos/toAdd/"

resolve = app.GetResolve()
if resolve:
    print("resolve load successful")
else:
    print("resolve load fail")
projectManager = resolve.GetProjectManager()
project = projectManager.GetCurrentProject()
timeline=project.GetCurrentTimeline()
recordFrame = TimecodeToFrame(timeline.GetCurrentTimecode(),round(project.GetSetting("timelineFrameRate")))
mediaPool = project.GetMediaPool()
addedMediaPoolItemList = mediaPool.ImportMedia(toAddFilePath)
if not addedMediaPoolItemList:
    print("video load fail,check file path")
for index,value in enumerate(addedMediaPoolItemList):
    #print(item.GetClipProperty())
    mediaPool.AppendToTimeline([{"mediaPoolItem":value,
                                "mediaType":1,
                                "trackIndex":index+1,
                                "recordFrame":recordFrame
                                }])
    # 添加具体逻辑，比如是否导入时当前时间线后面所有的timeline item都要往后移，看具体要求
