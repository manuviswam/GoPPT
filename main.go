package main

import (
	"fmt"
	"os"
	"path/filepath"

	z "github.com/manuviswam/GoPPT/zipper"
	m "github.com/manuviswam/GoPPT/model"
	fo "github.com/manuviswam/GoPPT/fileops"
)

const (
	mediaFolder  = "media"
	pptFolder = "ppt"
	templateImgName = "image1.jpg"
	slideFolder = "slides"
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
	}

	slidePath := filepath.Join(tempFolder,pptFolder,slideFolder,templateSlideName)

	replacements := make([]m.SlideReplacement,1)
	replacements = append(replacements,m.SlideReplacement{ PlaceHolder: "#title#", Replacement: "This is replaced title"})
	replacements = append(replacements,m.SlideReplacement{ PlaceHolder: "#footer#", Replacement: "This is replaced footer"})

	err = fo.ReplaceTextInFile(getAbsolutePath(slidePath),replacements)
	if err != nil {
		fmt.Println("Error while content replacing",err)
	}

	targetImgPath := filepath.Join(tempFolder,pptFolder,mediaFolder,templateImgName)

	err = fo.CopyFile(getAbsolutePath(newImage),getAbsolutePath(targetImgPath))
	if err != nil {
		fmt.Println("Error while image replacing",err)
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


