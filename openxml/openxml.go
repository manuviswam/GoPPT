package openxml

import (
	"path/filepath"

	uuid "github.com/satori/go.uuid"
	c "github.com/manuviswam/GoPPT/constants"
	fo "github.com/manuviswam/GoPPT/fileops"
	"fmt"
	"io/ioutil"
	"regexp"
)

const (
	relsExtension = ".rels"
	presentationRelFile = "presentation.xml.rels"
	presentationFile = "presentation.xml"
	relationshipXmlNode = `<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/%s" />`
	presentationSlIdNode = `<p:sldId id="%d" r:id="%s" />`
	relationshipsRegex = `(<Relationships xmlns="http:\/\/schemas\.openxmlformats\.org\/package\/2006\/relationships">)(?s)(.*)(<\/Relationships>)`
	presentationSlIdLstRegex = `(<p:sldIdLst>)(?s)(.*)(<\/p:sldIdLst>)`
	contentTypeRegex = `(<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">)(?s)(.*)(</Types>)`
	contentTypeNode = `<Override ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml" PartName="/ppt/slides/%s"/>`
)

func DuplicateSlide(pptRootPath, slideName string)(newSlideName string, err error){
	newSlideName = uuid.NewV4().String() + ".xml"
	newSlideName = "slilde2.xml"

	slidePath := filepath.Join(pptRootPath, c.PPTFolder, c.SlideFolder, slideName)
	newSlidePath := filepath.Join(pptRootPath, c.PPTFolder, c.SlideFolder, newSlideName)

	err = fo.CopyFile(slidePath, newSlidePath)
	if err != nil {
		return "", err
	}

	sourceSlideRelsPath := filepath.Join(pptRootPath, c.PPTFolder, c.SlideFolder, c.RelsFolder, slideName + relsExtension)
	newSlideRelsPath := filepath.Join(pptRootPath, c.PPTFolder, c.SlideFolder, c.RelsFolder, newSlideName + relsExtension)

	err = fo.CopyFile(sourceSlideRelsPath, newSlideRelsPath)
	if err != nil {
		return "", err
	}

	rid, err := createRelations(pptRootPath, newSlideName)
	if err != nil {
		return "", err
	}

	err = addSlideInPresentation(pptRootPath, rid)
	if err != nil {
		return "", err
	}

	return newSlideName, addSlideContentTypeInContentTypes(pptRootPath, newSlideName)
}

func createRelations(pptRoot, slideName string)(string, error) {
	presentationRelPath := filepath.Join(pptRoot, c.PPTFolder, c.RelsFolder, presentationRelFile)
	relationId := "rid100" //todo
	newRelationNode := fmt.Sprintf(relationshipXmlNode, relationId, slideName)

	return relationId, replaceInXmlFile(presentationRelPath, newRelationNode, relationshipsRegex)
}

func addSlideInPresentation(pptRoot, rid string)error {
	presentationXmlPath := filepath.Join(pptRoot, c.PPTFolder, presentationFile)
	slId := 300 //todo
	newSlIdNode := fmt.Sprintf(presentationSlIdNode, slId, rid)

	return replaceInXmlFile(presentationXmlPath, newSlIdNode, presentationSlIdLstRegex)
}

func addSlideContentTypeInContentTypes(pptRoot, newSlideName string)error {
	contentTypeXmlPath := filepath.Join(pptRoot, "[Content_Types].xml")
	newContentTypeNode := fmt.Sprintf(contentTypeNode, newSlideName)

	return replaceInXmlFile(contentTypeXmlPath, newContentTypeNode, contentTypeRegex)
}

func replaceInXmlFile(xmlPath, newNode, regexString string) error {
	presentationRelContentBytes, err := ioutil.ReadFile(xmlPath)
	if err != nil {
		return err
	}

	presentationRelContent := string(presentationRelContentBytes)

	re := regexp.MustCompile(regexString)

	relationships := re.FindStringSubmatch(presentationRelContent)[2]

	relationships += newNode

	presentationRelContent = re.ReplaceAllString(presentationRelContent, "$1" + relationships + "$3")

	return ioutil.WriteFile(xmlPath, []byte(presentationRelContent), 0755)
}
