package path

import (
	"fmt"
	"os"
	"os/exec"
	"path"
	"path/filepath"
)

func GetCurrentDirectory() string {
	file, _ := exec.LookPath(os.Args[0])
	return path.Dir(file)
}

func GetBaseFile(filepath string) string {
	return path.Base(filepath)
}

func GetBaseDir(filepath string) string {
	return path.Dir(filepath)
}

func GetFilelist(path string) []os.FileInfo {
	filelist := make([]os.FileInfo, 0)
	err := filepath.Walk(path, func(path string, f os.FileInfo, err error) error {
		if f == nil {
			return err
		}
		if f.IsDir() {
			return nil
		}

		filelist = append(filelist, f)
		return nil
	})

	if err != nil {
		fmt.Printf("filepath.Walk() returned %v\n", err)
	}

	return filelist
}

func WriteFile(file, content string) bool {
	f, err := os.Create(file)
	if err != nil {
		panic(err)
	}
	defer f.Close()

	f.WriteString(content)
	return true
}
