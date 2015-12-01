package main

import (
	"archive/zip"
	"fmt"
	"os"
	"path/filepath"
	"io"
	"strings"
	"io/ioutil"
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

	err := unzip(getAbsolutePath(sourcePPT), getAbsolutePath(tempFolder))
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

	err = zipit(getAbsolutePath(tempFolder), getAbsolutePath(targetPPT))

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

func unzip(archive, target string) error {
	reader, err := zip.OpenReader(archive)
	if err != nil {
		return err
	}

	if err := os.MkdirAll(target, 0755); err != nil {
		return err
	}

	for _, file := range reader.File {
		path := filepath.Join(target, file.Name)

		if file.FileInfo().IsDir() {
			os.MkdirAll(path, file.Mode())
			continue
		}

		os.MkdirAll(strings.Split(path,file.FileInfo().Name())[0], file.Mode())

		fileReader, err := file.Open()
		if err != nil {
			return err
		}
		defer fileReader.Close()

		targetFile, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, file.Mode())
		if err != nil {
			return err
		}
		defer targetFile.Close()

		if _, err := io.Copy(targetFile, fileReader); err != nil {
			return err
		}
	}

	return nil
}

func zipit(source, target string) error {
	zipfile, err := os.Create(target)
	if err != nil {
		return err
	}
	defer zipfile.Close()

	archive := zip.NewWriter(zipfile)
	defer archive.Close()

	info, err := os.Stat(source)
	if err != nil {
		return nil
	}

	var baseDir string
	if info.IsDir() {
		baseDir = filepath.Base(source)
	}
	skipOnce := true
	filepath.Walk(source, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if skipOnce {
			skipOnce = false
			return nil
		}

		header, err := zip.FileInfoHeader(info)
		if err != nil {
			return err
		}

		if baseDir != "" {
			header.Name = strings.TrimPrefix(path, source + "\\")
		}

		if info.IsDir() {
			header.Name += "/"
		} else {
			header.Method = zip.Deflate
		}

		writer, err := archive.CreateHeader(header)
		if err != nil {
			return err
		}

		if info.IsDir() {
			return nil
		}

		file, err := os.Open(path)
		if err != nil {
			return err
		}
		defer file.Close()
		_, err = io.Copy(writer, file)
		return err
	})

	return err
}
