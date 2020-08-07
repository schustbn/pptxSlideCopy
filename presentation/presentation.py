import zipHandler.zipHandler as zh
import relationships.relationship as r
import typeConfigs.typeConfigs as tc
import slide.slide as sl


class Presentation:
    """Presentation Class."""

    def __init__(self, file=None):
        """Init."""
        self.relationshipsLoaded = False
        if file is not None:
            self.readFromFile(file)

    def readFromFile(self, pptFilePath: str):
        """Read from File => Presentation."""
        self.presentationPath = pptFilePath
        self.presentationXml = zh.getXmlFileFromZip(
            "ppt/presentation.xml", pptFilePath
        )

        self.relations = list()

        # Get Relations and create relations
        for rel in zh.getXmlFileFromZip(
            "ppt/_rels/presentation.xml.rels", pptFilePath
        ):
            relation = r.Relationship(
                rel.get("Id"),
                rel.get("Type"),
                rel.get("Target"),
                rel.get("TargetMode"),
            )
            self.relations.append(relation)

        # Get Slides and create slides
        self.slides = list()
        for element in self.presentationXml:
            if element.tag == tc.slidelist_tag:
                for slide_element in element:
                    slide_filepath = (
                        "ppt/"
                        + self.getRelations(
                            slide_element.get(tc.relationship_id_tag)
                        ).target
                    )
                    slideNumber = int(
                        slide_filepath.partition("s/slide")[2].partition(".")[
                            0
                        ]
                    )
                    slide = sl.Slide(
                        id=element.get("Id"),
                        rId=element.get(tc.relationship_id_tag),
                        slideNo=slideNumber,
                        slidePath=slide_filepath,
                        presentationPath=self.presentationPath,
                    )
                    self.slides.append(slide)

        # get SlideLayouts and create slideLayouts
        self.slideMasters = list()
        for rel in zh.getXmlFileFromZip(
            "ppt/slideMasters/_rels/slideMaster1.xml.rels", pptFilePath
        ):
            if rel.get("Type") == tc.slideLayout_type:
                relation = r.Relationship(
                    rel.get("Id"),
                    rel.get("Type"),
                    rel.get("Target"),
                    rel.get("TargetMode"),
                )
                # print('Pfad: ', 'ppt' + relation.target.partition('..')[2])
                for elem in zh.getXmlFileFromZip(
                    "ppt" + relation.target.partition("..")[2], pptFilePath
                ):
                    if elem.tag == tc.slideLayout_tag:
                        # print(elem)
                        relation.addName(elem.get("name"))
                self.slideMasters.append(relation)

        # get HandoutsMaster and create handoutsMaster

        # get NotesMaster and create notesMaster

        # get LayoutMaster and create layoutMaster

    def getRelations(self, rId=None, relType=None):
        """Get Relations for Presentation."""
        # print('getRelations - params: ', rId, self.relations)
        rels = list()

        for element in self.relations:
            if rId is not None:
                if len(rId) > 0:
                    if element.id == rId:
                        return element
            elif relType is not None:
                if element.type == relType:
                    return element
            else:
                return self.relations
        return rels

    def getSlides(self):
        """Sldies."""
        return self.slides

    def getSlide(self, slideNo: int = None, rId: int = None) -> sl.Slide:
        """Returns one slide if present in the presentation."""
        slide_to_return = None
        if slideNo is not None:
            for slide in self.slides:
                if slide.slideNo == slideNo:
                    slide_to_return = slide
        elif rId is not None:
            for slide in self.slides:
                if slide.rId == rId:
                    slide_to_return = slide
        return slide_to_return

    def getSlideMasters(self):
        """SlideLayout."""
        return self.slideMasters

    def getSlideMasterAndIdForSlideMasterName(self, slideMastertName: str):
        """SlideMaster Path and ID."""
        slideMasters = self.getSlideMasters()
        for layout in slideMasters:
            if layout.name == slideMastertName:
                return layout

        return None

    def copySlideFromPresentation(
        self,
        source_presentation,
        slide_no: int = None,
        rId: int = None,
        new_position: int = None,
    ):
        """Copies the mentioned Slide from current Presentation to targetPresentation, including adapting SlideMaster, Images, pictures, media nad links.
        If position is empty (NOne) slide is appended at the End of the targetPresentation"""

        # copy Slide itself
        slide_to_copy = source_presentation.getSlide(slide_no, rId)
        if slide_to_copy is not None:
            slide_files = slide_to_copy.getSlideFileWithRelationFiles()
            for file in slide_files:
                print(file)
                if (
                    file["elemType"] == "slide"
                ):  # process the slide itself => copy and update the rIds => needeD???
                    # copy slide
                    zh.copyFile(
                        source_presentation.presentationPath,
                        self.presentationPath,
                        file["elemPath"],
                    )

                    # add slide to presentation.rels

                    # add slide to presentation.xml

                elif (
                    file["elemType"] == "rels"
                ):  # process the rels file => copy and update the rId/paths
                    # copy rels file
                    zh.copyFile(
                        source_presentation.presentationPath,
                        self.presentationPath,
                        file["elemPath"],
                    )

                    # update rId

                    # updateSlideLayout

                # slideLayout rausfiltern!!!
                elif (
                    file["elemType"] != tc.relTypeSlideLayout
                ):  # process all related files => copy and update rels

                    # copy file
                    zh.copyFile(
                        source_presentation.presentationPath,
                        self.presentationPath,
                        file["elemPath"],
                    )

                    # update rels file with new filename

        # rename all slide-files and slide_rels files
        # create rId for new Slide and update
        # add slide to targetpresentation.xml and targetPresentation_rels.xml
        # copy media and rename it
        # copy/create slide_rels.xml and reflect new media names

    def getSlideWithAllRels(self, slideNo: int = None, rId: int = None):
        slide = self.getSlide(slideNo, rId)
        slide.getRelations()
        return slide

    def reIndexSlides(self):
        """Re-Index the Slides."""
        pass
