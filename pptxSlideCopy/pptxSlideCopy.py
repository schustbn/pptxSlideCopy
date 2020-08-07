import presentation.presentation as pres
import zipfile as zf
import zipHandler.zipHandler as zh


sourecFile = "/Users/bjoernschuster/Downloads/PPTmerge/simpel.pptx"
targetFile = "/Users/bjoernschuster/Downloads/PPTmerge/master.pptx"


target_presentation = pres.Presentation(targetFile)
source_presentation = pres.Presentation(sourecFile)


for slide in source_presentation.slides:
    rels = slide.getRelations()
    notes = slide.getNotes()
    #    print('Notes: ', notes)
    slideLayout = slide.getSlideLayout()
    #    print("slide Layout: ", slideLayout)
    newLayout = target_presentation.getSlideMasterAndIdForSlideMasterName(
        slideLayout
    )

    #    print("new layout: ", newLayout.target)

    relFiles = slide.getSlideFileWithRelationFiles()
    # print("files: ", relFiles)

    print(slide.slidePath)
    if notes:
        pass
        # print(notes)
    for rel in rels:
        # print(rel.id, rel.target, '    Type: ', rel.type)
        pass


def getFileFromZip(file_path: str, zipfile_path: str) -> bytes:
    """Get File From ZIP."""
    zipfile = zf.ZipFile(zipfile_path)
    file = zipfile.read(file_path)
    return file


def addFileToZip(file_path: str, zipfile_path: str, file: bytes):
    zipfile = zf.ZipFile(zipfile_path, mode="a")
    zipfile.writestr(file_path, file)


filepath = "ppt/media/media1.mp4"
file = getFileFromZip(filepath, sourecFile)


# test = zh.copyFile(sourecFile, targetFile, filepath)
# print(test)


# test = zh.getNextFileNoInFilePath(filepath, zh.getContentsOfZipfileDirectory(targetFile, 'ppt/media'))
# print(test)

target_presentation.copySlideFromPresentation(source_presentation, 1)
