package main

import (
	"baliance.com/gooxml/document"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"strings"
)

func GetFileName() {
	if dir, err := os.Getwd(); err == nil {
		var files []string
		if ok := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
			files = append(files, path)
			return nil
		}); ok != nil {
			fmt.Println(ok)
		}
		for _, file := range files {
			fileName := filepath.Base(file)
			switch path.Ext(fileName) {
			case ".doc" , ".docx":
				saveFile(strings.Replace(fileName, path.Ext(fileName), ".txt", -1), makeDocx(fileName))
			default:
				fmt.Println(fileName, "不是docx文件，跳过处理！")

			}

		}
	} else {
		fmt.Println(err)
	}
}

func saveFile(fileName, content string) {
	content = strings.Replace(content, "\n　　\n", "\n", -1)
	fmt.Println("TextFile/"+fileName)
	if ok := ioutil.WriteFile("TextFile/"+fileName, []byte(content), 0644); ok != nil {
		log.Fatalf("error writing file: %s", ok)
	} else {
		fmt.Println("文件" + fileName + "已保存")
	}
}

func makeDocx(fileName string) string {
	var content string
	fmt.Println(fileName)
	if doc, err := document.Open(fileName); err != nil {
		log.Fatalf("error opening document: %s", err)
	} else {
		//doc.Paragraphs() 得到包含文档所有的段落的切片
		//run为每个段落相同格式的文字组成的片段
		for _, para := range doc.Paragraphs() {
			for _, run := range para.Runs() {
				content += run.Text()
			}
			content += "\n　　"
		}
		return content
	}
	return ""
}

func main() {
	if _, err := os.Stat("./TextFile"); err != nil {
		if err = os.Mkdir("./TextFile", 0777); err != nil {
			log.Fatalf("error creating directory: %s", err)
		}
	} else {
		fmt.Println("TextFile 目录已存在, 不再创建")
	}
	GetFileName()
}
