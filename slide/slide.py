import zipHandler.zipHandler as zh
import relationships.relationship as r
import typeConfigs.typeConfigs as tc


class Slide:
    """Class Docstring."""

    def __init__(
        self,
        id: int,
        rId: str,
        slideNo: int,
        slidePath: str,
        presentationPath: str,
    ):
        """Init."""
        self.id = id
        self.rId = rId
        self.slideNo = slideNo
        self.slidePath = slidePath
        self.presentationPath = presentationPath

    def getRelations(self):
        """Relations."""
        # print('Slide: getRelations')
        self.relations = list()

        slideRelsPath = (
            "ppt/slides/_rels/"
            + self.slidePath.partition("slides/")[2]
            + ".rels"
        )
        self.slideRelsPath = slideRelsPath
        slideRels = zh.getXmlFileFromZip(slideRelsPath, self.presentationPath)

        for element in slideRels:

            rel = r.Relationship(
                element.get("Id"),
                element.get("Type"),
                element.get("Target"),
                element.get("TargetMode"),
            )

            self.relations.append(rel)
        return self.relations

    def notesFilePath(self) -> str:
        """NotesFile."""
        path = None
        if not self.relations:
            self.getRelations()
        for rel in self.relations:
            if rel.type == tc.relTypeNotes:
                path = rel.target
                path = path.replace("..", "ppt")
        return path

    def getNotes(self):
        """Notes."""
        # print('getNotes')
        notesPath = self.notesFilePath()
        noteBody = None
        if notesPath is not None:
            notesXml = zh.getXmlFileFromZip(notesPath, self.presentationPath)
            noteBody = notesXml.findall(".//" + tc.slideNotesBody_tag)
            if len(noteBody) > 0:
                self.noteXml = noteBody
        return noteBody

    def getSlideFile(self):
        """SlideFile."""
        return zh.getXmlFileFromZip(self.slidePath, self.presentationPath)

    def getSlideLayout(self):
        """SlideLayout."""
        #        print('getSlideLayout')
        if not self.relations:
            self.getRelations()
        for rel in self.relations:
            if rel.type == tc.relTypeSlideLayout:
                for elem in zh.getXmlFileFromZip(
                    "ppt" + rel.target.partition("..")[2],
                    self.presentationPath,
                ):
                    if elem.tag == tc.slideLayout_tag:
                        rel.addName(elem.get("name"))
                        self.slideLayoutName = elem.get("name")
        return self.slideLayoutName

    def getSlideFileWithRelationFiles(self):
        """SlideFile."""
        self.files = list()
        slidepath = {"elemPath": self.slidePath, "elemType": "slide"}
        relspath = {"elemPath": self.slideRelsPath, "elemType": "rels"}
        self.files.append(slidepath)
        self.files.append(relspath)
        if not self.relations:
            self.getRelations()
        for rel in self.relations:
            if rel.targetMode != "External":
                filePath = "ppt" + rel.target.partition("..")[2]
                # print("filePath: ", filePath, self.slidePath, rel)
                fileElem = {
                    "elemPath": filePath,
                    #                   'file': getXmlFileFromZip(filePath, self.presentationPath),
                    "elemType": rel.type,
                }
                self.files.append(fileElem)
        return self.files
