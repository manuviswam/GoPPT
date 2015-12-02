package fileops
import (
	"io"
	"os"
	"io/ioutil"
	"strings"
)

func ReplaceTextInFile(filename string, replacements []SlideReplacement)error {
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

func CopyFile(sourceImageFilename, targetImageFilename string) error {
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
