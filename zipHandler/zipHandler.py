import zipfile as zf
import lxml.etree
import re


def getXmlFileFromZip(file_path: str, zipfiile_path: str):
    """Get XML File from ZIP."""
    zipfile = zf.ZipFile(zipfiile_path)
    file_string = zipfile.read(file_path)
    file_xml = lxml.etree.fromstring(file_string)
    return file_xml


def getFileFromZip(file_path: str, zipfile_path: str) -> bytes:
    """Get File From ZIP."""
    zipfile = zf.ZipFile(zipfile_path)
    file = zipfile.read(file_path)
    return file


def addFileToZip(file_path: str, zipfile_path: str, file: bytes):
    zipfile = zf.ZipFile(zipfile_path, mode="a")
    zipfile.write(file, arcname=file_path)


def getContentsOfZipfileDirectory(zip_filepath: str, directory: str) -> list:
    zipfile = zf.ZipFile(zip_filepath)
    resultList = []
    for elem in zipfile.namelist():
        if elem.startswith(directory):
            resultList.append(elem)

    return resultList


def getDirectoryFromPath(path: str) -> str:
    """Retrieves the directory string from a path string."""
    path_temp = path.rpartition("/")
    new_path = path_temp[0] + path_temp[1]
    return new_path


def changeFileNoInFilePath(path: str, fileNo: int) -> str:
    """replaces the number in the path with the given number."""

    separator = r"[0-9]+\."
    splitted_path = re.split(separator, path, 1)
    new_path = splitted_path[0] + str(fileNo) + "." + splitted_path[1]
    return new_path


def getNextFileNoInFilePath(path: str, directory_list: list) -> int:
    print("getNextFileNoInFilePath : ", path, directory_list)
    path_type = getTypeOfFilePath(path)

    matched_paths = []

    for file in directory_list:
        print(path_type + " = " + getTypeOfFilePath(file))
        if path_type == getTypeOfFilePath(file):
            matched_paths.append(file)
    print(matched_paths)

    nextNo = len(matched_paths) + 1

    if len(matched_paths) < 1:
        nextNo = 1

    return nextNo


def getTypeOfFilePath(filepath: str) -> str:
    separator = "[a-z]+[0-9]+"
    match = re.search(separator, filepath)
    type = ""
    if match:
        type = re.search("[a-z]+", match.group())
    return type.group()


def copyFile(
    source_filepath: str, target_filepath: str, file_filepath: str
) -> str:
    source = zf.ZipFile(source_filepath, "r")
    source_file_content = source.read(file_filepath)
    source.close()

    target = zf.ZipFile(
        target_filepath, "a"
    )  # das 'a' ist wichtig, sonst wir der Inhalt des Zipfile überschrieben

    directoryFiles = getContentsOfZipfileDirectory(
        target_filepath, getDirectoryFromPath(file_filepath)
    )

    file_filepath = changeFileNoInFilePath(
        file_filepath, getNextFileNoInFilePath(file_filepath, directoryFiles)
    )

    target.writestr(file_filepath, source_file_content)

    target.close()
    return file_filepath


# def writeSlidesIntoPresentation(slides, target_filepath):
#     target = zf.ZipFile(
#         target_filepath, "a"
#     )  # das 'a' ist wichtig, sonst wird der Inhalt des Zipfile überschrieben
#     presentation_string = target.read("ppt/presentation.xml")
#     presentation_xml = lxml.etree.fromstring(presentation_string)
#
#     presentation_rels_string = target.read("ppt/_rels/presentation.xml.rels")
#     presentation_rels_xml = lxml.etree.fromstring(presentation_rels_string)
#
#     # rels für slides löschen
#     for rel in presentation_rels_xml:
#         if (
#             rel.get("Type")
#             == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
#         ):
#             presentation_rels_xml.remove(rel)
#     print(presentation_rels_xml)
#
#     # find elem => elem.getparent().remove(elem)
#     index = 0
#     for elem in presentation_xml:
#         if (
#             elem.tag
#             == "{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst"
#         ):
#             for slide_element in elem:
#                 elem.remove(slide_element)
#             for slide in slides:
#                 slide_child = lxml.etree.Element(
#                     "{http://schemas.openxmlformats.org/drawingml/2006/main}sldId"
#                 )
#                 slide_child.set("Id", str(slide["slide_id"]))
#                 slide_child.set(
#                     "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id",
#                     str(slide["slide_r_id"]),
#                 )
#                 elem.insert(index, slide_child)
#
#                 rel_child = lxml.etree.Element("Relationship")
#                 rel_child.set("Id", str(slide["slide_r_id"]))
#                 rel_child.set(
#                     "Type",
#                     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
#                 )
#                 rel_child.set("Target", slide["filepath"])
#                 presentation_rels_xml.insert(index, rel_child)
#
#                 index += 1
#
#     # slide_rels hinzufügen
#     writeSlideRels(slides, target_filepath)
#
#     target.writestr(
#         "ppt/presentation.xml", lxml.etree.tostring(presentation_xml)
#     )
#     target.writestr(
#         "ppt/_rels/presentation.xml.rels",
#         lxml.etree.tostring(presentation_rels_xml),
#     )


# def writeSlideRels(slides, target_filepath):
#     target = zf.ZipFile(
#         target_filepath, "a"
#     )  # das 'a' ist wichtig, sonst wir der Inhalt des Zipfile überschrieben
#     presentation_xml = lxml.etree.fromstring(
#         target.read("ppt/presentation.xml")
#     )
#
#     rels_initial_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
#     rels_xml = lxml.etree.fromstring(rels_initial_string)
