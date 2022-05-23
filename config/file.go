package config

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"strings"
)

func HomeDir() string {
	if dir, err := os.Getwd(); err == nil {
		return dir
	} else {
		fmt.Println(err)
	}
	return ""
}

func FileNameList(filePath string) []string {
	var files []string
	if ok := filepath.Walk(path.Join(HomeDir(), filePath),
		func(path string, info os.FileInfo, err error) error {
			files = append(files, path)
			return nil
		},
	); ok == nil {
		return files
	}
	return nil
}

func SaveFile(fileName, content, filePath string) {
	NewName := strings.Replace(fileName, ".docx", ".txt", -1)
	content = strings.Replace(content, "\n　　\n", "\n", -1)
	if ok := ioutil.WriteFile(filePath+NewName, []byte(content), 0644); ok != nil {
		log.Fatalf("error writing file: %s", ok)
	}
}

func MkdirFile(filepath string) {
	if _, err := os.Stat(filepath); err != nil {
		if err = os.Mkdir(filepath, 0777); err != nil {
			log.Fatalf("error creating directory: %s", err)
		}
	}
}
