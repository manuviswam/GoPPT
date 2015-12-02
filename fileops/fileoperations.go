package fileops
import (
	"io"
	"os"
	"io/ioutil"
	"strings"

	m "github.com/manuviswam/GoPPT/model"
)

func ReplaceTextInFile(filename string, replacements []m.SlideReplacement)error {
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

func CopyFile(sourceFilename, targetFilename string) error {
	r, err := os.Open(sourceFilename)
	if err != nil {
		return err
	}
	defer r.Close()

	w, err := os.Create(targetFilename)
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
