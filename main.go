package main

import (
	"fmt"
	"os"
	"path/filepath"

	z "github.com/manuviswam/GoPPT/zipper"
	m "github.com/manuviswam/GoPPT/model"
	fo "github.com/manuviswam/GoPPT/fileops"
	c "github.com/manuviswam/GoPPT/constants"
	"github.com/manuviswam/GoPPT/openxml"
	"github.com/getgauge/gauge/util"
)

const (
	templateImgName = "image1.jpg"
	templateSlideName = "slide1.xml"
)

func main() {
	sourcePPT := "Template.pptx"
	tempFolder := "tmp"
	targetPPT := "new.pptx"
	newImage := "img.jpg"

	err := z.Unzip(getAbsolutePath(sourcePPT), getAbsolutePath(tempFolder))
	if err != nil {
		fmt.Println("Error unziping : ", err)
		return
	}
	defer util.Remove(getAbsolutePath(tempFolder))

	//duplicate slide
	newSlideName, err := openxml.DuplicateSlide(getAbsolutePath(tempFolder), templateSlideName)
	if err != nil {
		fmt.Println("Error while duplicationg slide : ",err)
		return
	}

	slidePath := filepath.Join(tempFolder,c.PPTFolder,c.SlideFolder,newSlideName)

	replacements := make([]m.SlideReplacement,1)
	replacements = append(replacements,m.SlideReplacement{ PlaceHolder: "#title#", Replacement: "This is replaced title"})
	replacements = append(replacements,m.SlideReplacement{ PlaceHolder: "#footer#", Replacement: "This is replaced footer"})


	//replace text placeholders
	err = fo.ReplaceTextInFile(getAbsolutePath(slidePath),replacements)
	if err != nil {
		fmt.Println("Error while content replacing",err)
		return
	}

	targetImgPath := filepath.Join(tempFolder,c.PPTFolder,c.MediaFolder,templateImgName)

	//replace image placeholder
	err = fo.CopyFile(getAbsolutePath(newImage),getAbsolutePath(targetImgPath))
	if err != nil {
		fmt.Println("Error while image replacing",err)
		return
	}


	err = z.Zipit(getAbsolutePath(tempFolder), getAbsolutePath(targetPPT))

	if err != nil {
		fmt.Println("Error ziping : ", err)
	}
}

func getAbsolutePath(relPath string)string  {
	wd,_ := os.Getwd()
	return filepath.Join(wd, relPath)
}


