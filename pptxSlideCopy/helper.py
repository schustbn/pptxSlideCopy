import zipfile as zf
import lxml.etree
import xmltodict
import pprint
import json


slidelist_tag = '{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst'
relationship_id_tag = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
slide_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
slideLayout_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'

slideNotesBody_tag = '{http://schemas.openxmlformats.org/presentationml/2006/main}txBody'

sourecFile='/Users/bjoernschuster/Downloads/PPTmerge/simpel.pptx'
targetFile='/Users/bjoernschuster/Downloads/PPTmerge/master.pptx'

slideLayout_tag = '{http://schemas.openxmlformats.org/presentationml/2006/main}cSld'


# Slide Rel-Types:
relTypeHyperlink = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
relTypeNotes = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide'
relTypeSlideLayout = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
relTypeImage = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
relTypeVideo = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'
relTypeMedia = 'http://schemas.microsoft.com/office/2007/relationships/media'


class Slide:
    def __init__(self, id: int, rId: str, slideNo: int, slidePath: str, presentationPath: str):
        self.id = id
        self.rId = rId
        self.slideNo = slideNo
        self.slidePath = slidePath
        self.presentationPath = presentationPath

    def getRelations(self):
        #print('Slide: getRelations')
        self.relations = list()
        pp = pprint.PrettyPrinter(indent=4)

        slideRelsPath = 'ppt/slides/_rels/' + self.slidePath.partition('slides/')[2] + '.rels'

        slideRels = getXmlFileFromZip(slideRelsPath, self.presentationPath)

        for element in slideRels:

            rel = Relationship(element.get('Id'), element.get('Type'), element.get('Target'), element.get('TargetMode'))

            self.relations.append(rel)
        return self.relations

    def getNotes(self):
        #print('getNotes')
        if not self.relations:
            self.getRelations()
        for rel in self.relations:
            if rel.type == relTypeNotes:
                path = rel.target
                path = path.replace('..', 'ppt')
                notesXml = getXmlFileFromZip(path, self.presentationPath)

                noteBody = notesXml.findall('.//' + slideNotesBody_tag)
                if len(noteBody) >0:
                    self.noteXml = noteBody
                    return self.noteXml

    def getSlideFile(self):
        return getXmlFileFromZip(self.slidePath, self.presentationPath)

    def getSlideLayout(self):
#        print('getSlideLayout')
        if not self.relations:
            self.getRelations()
        for rel in self.relations:
            if rel.type == relTypeSlideLayout:
                for elem in getXmlFileFromZip('ppt' + rel.target.partition('..')[2], self.presentationPath):
                    if elem.tag == slideLayout_tag:
                        rel.addName(elem.get('name'))
                        self.slideLayoutName = elem.get('name')
        return self.slideLayoutName

    def getSlideFileWithRelationFiles(self):
        self.files = list()
        if not self.relations:
            self.getRelations()
        for rel in self.relations:
            if rel.targetMode != 'External':
                filePath = 'ppt' + rel.target.partition('..')[2]
                print ('filePath: ', filePath, self.slidePath, rel)
                fileElem = {
                    'elemPath': filePath,
 #                   'file': getXmlFileFromZip(filePath, self.presentationPath),
                    'elemType': rel.type
                }
                self.files.append(fileElem)
        return self.files


class Presentation:
    def __init__(self):
        self.relationshipsLoaded = False

    def readFromFile(self, pptFilePath: str):
        self.presentationPath = pptFilePath
        self.presentationXml = getXmlFileFromZip('ppt/presentation.xml', pptFilePath)

        self.relations = list()

        # Get Relations and create relations
        for rel in getXmlFileFromZip('ppt/_rels/presentation.xml.rels', pptFilePath):
            relation = Relationship(rel.get('Id'), rel.get('Type'), rel.get('Target'), rel.get('TargetMode'))
            self.relations.append(relation)

        # Get Slides and create slides
        self.slides = list()
        for element in self.presentationXml:
            if element.tag == slidelist_tag:
                for slide_element in element:
                    slide_filepath = 'ppt/' + self.getRelations(slide_element.get(relationship_id_tag)).target
                    slideNumber = int(slide_filepath.partition('s/slide')[2].partition('.')[0])
                    slide = Slide(id=element.get('Id'), rId=element.get(relationship_id_tag), slideNo=slideNumber, slidePath=slide_filepath, presentationPath=self.presentationPath)
                    self.slides.append(slide)

        #get SlideLayouts and create slideLayouts
        self.slideMasters = list()
        for rel in getXmlFileFromZip('ppt/slideMasters/_rels/slideMaster1.xml.rels', pptFilePath):
            if rel.get('Type') == slideLayout_type:
                relation = Relationship(rel.get('Id'), rel.get('Type'), rel.get('Target'), rel.get('TargetMode'))
               # print('Pfad: ', 'ppt' + relation.target.partition('..')[2])
                for elem in getXmlFileFromZip('ppt' + relation.target.partition('..')[2], pptFilePath):
                    if elem.tag == slideLayout_tag:
                        #print(elem)
                        relation.addName(elem.get('name'))
                self.slideMasters.append(relation)

        #get HandoutsMaster and create handoutsMaster

        #get NotesMaster and create notesMaster

        #get LayoutMaster and create layoutMaster

    def getRelations(self, rId=''):
        #print('getRelations - params: ', rId)
        rels = list()
        pp = pprint.PrettyPrinter(indent=4)

        for element in self.relations:
            if rId != None:
                if len(rId) > 0:
                    if element.id == rId:
                        return element
            else:
                return self.relations
        return rels

    def getSlides(self):
        return self.slides

    def getSlideLayouts(self):
        pass
    def addSlide(self):
        pass
    def reIndexSlides(self):
        pass

class Relationship:
    def __init__(self, id:str, type:str, target:str, targetMode=''):
        self.id = id
        self.type = type
        self.target = target
        self.targetMode = targetMode
    def addName(self, name):
        #print ('Name: ' + name)
        self.name = name

def getXmlFileFromZip(file_path: str, zipfiile_path: str):
    zipfile = zf.ZipFile(zipfiile_path)
    file_string = zipfile.read(file_path)
    file_xml = lxml.etree.fromstring(file_string)
    return file_xml


# def getSlidesList(zipfile_path: str) -> list:
#     presentation_xml_path = 'ppt/presentation.xml'
#     presentation_rels_path = 'ppt/_rels/presentation.xml.rels'
#
#     slides_rels_path = 'ppt/slides/_rels/'
#
#     presentation_xml = getXmlFileFromZip(presentation_xml_path, zipfile_path)
#     slides = list()
#     for element in presentation_xml:
#         if element.tag == slidelist_tag:
#             for slide_element in element:
#                 slide_filepath = getRelsListForR_Id(document, presentation_rels_path,
#                                                     slide_element.get(relationship_id_tag))
#
#                 slide_rels_filepath = slides_rels_path + slide_filepath[0]['slide_filepath'].partition('/')[
#                     2] + '.rels'
#
#                 slide_rels = getRelsList(document, slide_rels_filepath)  # , ['image', 'video', 'media'])
#
#                 slide = {
#                     'slide_id': slide_element.get('id'),
#                     'slide_r_id': slide_element.get(relationship_id_tag),
#                     'filepath': slide_filepath[0]['slide_filepath'],
#                     'slide_rels_filepath': slide_rels_filepath
#                 }
#                 # if len(slide_rels) >0:
#                 slide['slide_rels'] = slide_rels
#                 slides.append(slide)
#     return slides
#
#
# def getRelsListForR_Id(document, rels_relative_filepath: str, r_id: str) -> list:
#     rels_string = document.read(rels_relative_filepath)
#     rels_xml = lxml.etree.fromstring(rels_string)
#
#     rels = list()
#
#     for element in rels_xml:
#         if element.get('Id') == r_id:
#             rel = {
#                 'r_id': element.get('Id'),
#                 'rel_type': element.get('Type'),
#                 'type': element.get('Type').rpartition('/')[2],
#                 'slide_filepath': element.get('Target')
#             }
#             rels.append(rel)
#     return rels
#
#
# def getRelsList(document, rels_relative_filepath, rel_type=None):
#     rels_string = document.read(rels_relative_filepath)
#     rels_xml = lxml.etree.fromstring(rels_string)
#
#     rels = list()
#
#     for element in rels_xml:
#         if len(rel_type) < 1:
#             rel = {
#                 'r_id': element.get('Id'),
#                 'rel_type': element.get('Type'),
#                 'type': element.get('Type').rpartition('/')[2],
#                 'slide_filepath': element.get('Target')
#             }
#             rels.append(rel)
#         else:
#             for elem in rel_type:
#                 if element.get('Type').endswith(elem):
#                     rel = {
#                         'r_id': element.get('Id'),
#                         'rel_type': element.get('Type'),
#                         'type': element.get('Type').rpartition('/')[2],
#                         'slide_filepath': element.get('Target')
#                     }
#                     rels.append(rel)
#     return rels
#
#
# def copyFile(source_filepath, target_filepath, file_filepath):
#     target = zf.ZipFile(target_filepath,
#                         'a')  # das 'a' ist wichtig, sonst wir der Inhalt des Zipfile überschrieben
#     source = zf.ZipFile(source_filepath, 'r')
#
#     path = zf.Path(target_filepath, 'ppt/slides/')
#     target_file_counter = 1
#     for n in path.iterdir():
#         if n.is_file():
#             target_file_counter += 1
#
#     file_id = file_filepath.partition('.')[0].rpartition('slide')[2]
#
#     target_file_filepath = file_filepath.replace(file_id, str(target_file_counter))
#
#     source_file_content = source.read(file_filepath)
#     target_result = target.writestr(target_file_filepath, source_file_content)
#
#     # target.close()
#     # source.close()
#     return target_file_filepath
#
#
# def addSlideToPresentation(slide_filepath, target_filepath, slide):
#     currentSlides = getSlidesList(target_filepath)
#
#     slide_filepath = slide_filepath.partition('/')[2]
#
#     newSlide = {
#         'slide_id': '504',
#         'slide_r_id': 'rId5',
#         'filepath': slide_filepath,
#         'slide_rels_filepath': 'ppt/slides/_rels/' + slide_filepath.partition('/')[2] + '.rels'
#     }
#
#     currentSlides.append(newSlide)
#
#     reIndexedSlides = reIndexPresentationSlides(currentSlides)
#
#     writeSlidesIntoPresentation(reIndexedSlides, target_filepath)
#
#
# def reIndexPresentationSlides(presentation_slides):
#     id = 9800
#     rId = 1
#
#     for slide in presentation_slides:
#         slide['old_r_id'] = slide['slide_r_id']
#         slide['slide_id'] = id
#         slide['slide_r_id'] = rId
#         id += 1
#         rId += 1
#     return presentation_slides
#
#
# def writeSlidesIntoPresentation(slides, target_filepath):
#     target = zf.ZipFile(target_filepath,
#                         'a')  # das 'a' ist wichtig, sonst wir der Inhalt des Zipfile überschrieben
#     presentation_string = target.read('ppt/presentation.xml')
#     presentation_xml = lxml.etree.fromstring(presentation_string)
#
#     presentation_rels_string = target.read('ppt/_rels/presentation.xml.rels')
#     presentation_rels_xml = lxml.etree.fromstring(presentation_rels_string)
#
#     # rels für slides löschen
#     for rel in presentation_rels_xml:
#         if rel.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide':
#             presentation_rels_xml.remove(rel)
#     print(presentation_rels_xml)
#
#     # find elem => elem.getparent().remove(elem)
#     index = 0
#     for elem in presentation_xml:
#         if elem.tag == '{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst':
#             for slide_element in elem:
#                 elem.remove(slide_element)
#             for slide in slides:
#                 slide_child = lxml.etree.Element('{http://schemas.openxmlformats.org/drawingml/2006/main}sldId')
#                 slide_child.set('Id', str(slide['slide_id']))
#                 slide_child.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id',
#                                 str(slide['slide_r_id']))
#                 elem.insert(index, slide_child)
#
#                 rel_child = lxml.etree.Element('Relationship')
#                 rel_child.set('Id', str(slide['slide_r_id']))
#                 rel_child.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
#                 rel_child.set('Target', slide['filepath'])
#                 presentation_rels_xml.insert(index, rel_child)
#
#                 index += 1
#
#     # slide_rels hinzufügen
#     writeSlideRels(slides, target_filepath)
#
#     target.writestr('ppt/presentation.xml', lxml.etree.tostring(presentation_xml))
#     target.writestr('ppt/_rels/presentation.xml.rels', lxml.etree.tostring(presentation_rels_xml))
#
#
# def writeSlideRels(slides, target_filepath):
#     target = zf.ZipFile(target_filepath,
#                         'a')  # das 'a' ist wichtig, sonst wir der Inhalt des Zipfile überschrieben
#     presentation_xml = lxml.etree.fromstring(target.read('ppt/presentation.xml'))
#
#     rels_initial_string = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
#     rels_xml = lxml.etree.fromstring(rels_initial_string)
#
# #   for slide in slides:


source = Presentation()
source.readFromFile(sourecFile)

target = Presentation()
target.readFromFile(targetFile)

for slide in source.slides:
    rels = slide.getRelations()
    notes = slide.getNotes()
    slideLayout = slide.getSlideLayout()
    relFiles = slide.getSlideFileWithRelationFiles()
    print('files: ', relFiles)

#    print(slide.slidePath)
    if notes:
        pass
        #print(notes)
    for rel in rels:
        #print(rel.id, rel.target, '    Type: ', rel.type)
        pass