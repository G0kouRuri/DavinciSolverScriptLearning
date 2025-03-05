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



resolve = app.GetResolve()
if resolve:
    print("resolve load successful")
else:
    print("resolve load fail")
projectManager = resolve.GetProjectManager()
project=projectManager.GetCurrentProject()
projectFps=round(project.GetSetting("timelineFrameRate"))
print("Project FPS:",projectFps)
timeline=project.GetCurrentTimeline()
if not timeline:
    print("No Timeline excist!")

timeCode=timeline.GetCurrentTimecode()
print("Current Time Code:",timeCode)

mediaPool=project.GetMediaPool()
if not mediaPool:
    print("Get Media Pool Fail!")

ItemsInCurrentTimeline =[]
for trackType in ["video","audio"]:
    for track_index in range(1,timeline.GetTrackCount(trackType)+1):
        for item in timeline.GetItemListInTrack(trackType,track_index):
            if (item.GetStart(True)<=TimecodeToFrame(timeCode,projectFps) and item.GetEnd()>=TimecodeToFrame(timeCode,projectFps)):
                ItemsInCurrentTimeline.append(item)
                print(item.GetName())
                print(item.GetStart(False))
                print(item.GetEnd(False))
                print(item.GetSourceStartFrame())
                print(item.GetSourceEndFrame())
                print(item.GetProperty())
                print(item.GetFusionCompCount())
                
if not ItemsInCurrentTimeline:
    print("No items found at current timecode!")

'''
# **导出逻辑**
for index, item in enumerate(ItemsInCurrentTimeline):
    clipName = item.GetName()
    exportPath = f"E:/czyapp/Davinci/Projects/Output_Videos/{clipName}.mp4"

    # 创建新的timeline用于导出
    timelineTmp=mediaPool.CreateEmptyTimeline(timelineTmp)

    # 设置渲染任务
    renderSettings = {
        "TargetDir": "E:/czyapp/Davinci/Projects/Output_Videos/",
        "CustomName": f"{clipName}",
        "Format": "mp4",
        "VideoCodec": "H264",
        "AudioCodec": "AAC",
        "ExportVideo": True,
        "ExportAudio": True,
    }

    # 创建新的渲染任务
    jobId = project.AddRenderJob(timelineTmp, renderSettings)
    if jobId:
        print(f"Exporting: {clipName} to {exportPath}")
    else:
        print(f"Failed to create render job for {clipName}")

# 开始渲染
project.StartRendering()
'''
