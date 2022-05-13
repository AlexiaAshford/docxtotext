package config

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path/filepath"
	"strings"
)

func FileNameList() []string {
	var files []string
	if dir, err := os.Getwd(); err == nil {
		if ok := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
			files = append(files, path)
			return nil
		}); ok == nil {
			return files
		}
	} else {
		fmt.Println(err)
	}
	return nil
}

func SaveFile(fileName, content string) {
	content = "　　" + strings.Replace(content, "\n　　\n", "\n", -1)
	if ok := ioutil.WriteFile("TextFile/"+fileName, []byte(content), 0644); ok != nil {
		log.Fatalf("error writing file: %s", ok)
	} else {
		fmt.Println("文件" + fileName + "处理完毕， 已保存在 TextFile 目录下")
	}
}

func MkdirFile(filepath string) {
	if _, err := os.Stat(filepath); err != nil {
		if err = os.Mkdir(filepath, 0777); err != nil {
			log.Fatalf("error creating directory: %s", err)
		}
	} else {
		fmt.Println(filepath, " 目录已存在, 不再创建")
	}
}
