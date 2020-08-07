"""Helper."""

# import zipHandler.zipHandler as zh
# import typeConfigs.typeConfigs as tc
# import lxml.etree


# def getSlidesList(zipfile_path: str) -> list:
#     presentation_xml_path = "ppt/presentation.xml"
#     presentation_rels_path = "ppt/_rels/presentation.xml.rels"
#
#     slides_rels_path = "ppt/slides/_rels/"
#
#     presentation_xml = zh.getXmlFileFromZip(presentation_xml_path, zipfile_path)
#     slides = list()
#     for element in presentation_xml:
#         if element.tag == tc.slidelist_tag:
#             for slide_element in element:
#                 slide_filepath = getRelsListForR_Id(
#                     zipfile_path,
#                     presentation_rels_path,
#                     slide_element.get(tc.relationship_id_tag),
#                 )
#
#                 slide_rels_filepath = (
#                     slides_rels_path
#                     + slide_filepath[0]["slide_filepath"].partition("/")[2]
#                     + ".rels"
#                 )
#
#                 slide_rels = getRelsList(
#                     zipfile_path, slide_rels_filepath
#                 )  # ,\['image', 'video', 'media'])
#
#                 slide = {
#                     "slide_id": slide_element.get("id"),
#                     "slide_r_id": slide_element.get(tc.relationship_id_tag),
#                     "filepath": slide_filepath[0]["slide_filepath"],
#                     "slide_rels_filepath": slide_rels_filepath,
#                     "slide_rels": slide_rels,
#                 }
#                 slides.append(slide)
#     return slides


# def getRelsListForR_Id(document, rels_relative_filepath: str, r_id: str) -> list:
#     rels_string = document.read(rels_relative_filepath)
#     rels_xml = lxml.etree.fromstring(rels_string)
#
#     rels = list()
#
#     for element in rels_xml:
#         if element.get("Id") == r_id:
#             rel = {
#                 "r_id": element.get("Id"),
#                 "rel_type": element.get("Type"),
#                 "type": element.get("Type").rpartition("/")[2],
#                 "slide_filepath": element.get("Target"),
#             }
#             rels.append(rel)
#     return rels


# def getRelsList(document, rels_relative_filepath, rel_type=None):
#     rels_string = document.read(rels_relative_filepath)
#     rels_xml = lxml.etree.fromstring(rels_string)
#
#     rels = list()
#
#     for element in rels_xml:
#         if len(rel_type) < 1:
#             rel = {
#                 "r_id": element.get("Id"),
#                 "rel_type": element.get("Type"),
#                 "type": element.get("Type").rpartition("/")[2],
#                 "slide_filepath": element.get("Target"),
#             }
#             rels.append(rel)
#         else:
#             for elem in rel_type:
#                 if element.get("Type").endswith(elem):
#                     rel = {
#                         "r_id": element.get("Id"),
#                         "rel_type": element.get("Type"),
#                         "type": element.get("Type").rpartition("/")[2],
#                         "slide_filepath": element.get("Target"),
#                     }
#                     rels.append(rel)
#     return rels


###### ZIP-File Handling


# def addSlideToPresentation(slide_filepath, target_filepath, slide):
#     currentSlides = getSlidesList(target_filepath)
#
#     slide_filepath = slide_filepath.partition("/")[2]
#
#     newSlide = {
#         "slide_id": "504",
#         "slide_r_id": "rId5",
#         "filepath": slide_filepath,
#         "slide_rels_filepath": "ppt/slides/_rels/"
#         + slide_filepath.partition("/")[2]
#         + ".rels",
#     }
#
#     currentSlides.append(newSlide)
#
#     reIndexedSlides = reIndexPresentationSlides(currentSlides)
#
#     zh.writeSlidesIntoPresentation(reIndexedSlides, target_filepath)


# def reIndexPresentationSlides(presentation_slides):
#     id = 9800
#     rId = 1
#
#     for slide in presentation_slides:
#         slide["old_r_id"] = slide["slide_r_id"]
#         slide["slide_id"] = id
#         slide["slide_r_id"] = rId
#         id += 1
#         rId += 1
#     return presentation_slides


#   for slide in slides:
