package main

import (
	"fmt"
	"os"
	"path/filepath"
	"io"
	"strings"
	"io/ioutil"

	z "github.com/manuviswam/GoPPT/zipper"
)

const (
	mediaFolder  = "media"
	pptFolder = "ppt"
	templateImgName = "image1.jpg"
	slideFolder = "slides"
	templateSlideName = "slide1.xml"
)

type SlideReplacement struct {
	PlaceHolder, Replacement string
}

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

	replacements := make([]SlideReplacement,1)
	replacements = append(replacements,SlideReplacement{ PlaceHolder: "#title#", Replacement: "This is replaced title"})
	replacements = append(replacements,SlideReplacement{ PlaceHolder: "#footer#", Replacement: "This is replaced footer"})

	err = replaceText(getAbsolutePath(slidePath),replacements)
	if err != nil {
		fmt.Println("Error while content replacing",err)
	}

	targetImgPath := filepath.Join(tempFolder,pptFolder,mediaFolder,templateImgName)

	err = replaceImage(getAbsolutePath(newImage),getAbsolutePath(targetImgPath))
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

func replaceText(filename string, replacements []SlideReplacement)error {
	contentBytes,err := ioutil.ReadFile(filename)
	if err != nil {
		return err
	}

	content := string(contentBytes)

	for _,replacement := range replacements  {
		content = strings.Replace(content, replacement.PlaceHolder, replacement.Replacement, -1)
	}

	err = ioutil.WriteFile(filename,[]byte(content), 0755)
	if err != nil {
		return err
	}
	return nil
}

func replaceImage(sourceImageFilename, targetImageFilename string) error {
	r, err := os.Open(sourceImageFilename)
	if err != nil {
		return err
	}
	defer r.Close()

	w, err := os.Create(targetImageFilename)
	if err != nil {
		return err
	}
	defer w.Close()

	_, err = io.Copy(w, r)
	if err != nil {
		return err
	}
	return nil
}
