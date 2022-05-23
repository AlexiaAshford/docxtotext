package main

import (
	"baliance.com/gooxml/document"
	"doc/config"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path"
	"path/filepath"
	"sync"
)

var Vars *config.ClassConfig

func getDocxInformation(fileName string, index int, ch chan struct{}, wg *sync.WaitGroup) {
	var content string //doc.Paragraphs() 得到包含文档所有的段落的切片
	if doc, err := document.Open(Vars.DocxFileName + fileName); err == nil {
		for _, para := range doc.Paragraphs() {
			//run为每个段落相同格式的文字组成的片段
			content += "\n　　"
			for _, run := range para.Runs() {
				content += run.Text()
			}
		}
	} else {
		fmt.Println("fileName", err)
	}
	if content != "" {
		config.SaveFile(fileName, content, Vars.TextFileName)
		fmt.Println("No:", index, "\t文件", fileName, "处理完成")
		Vars.DelFileList = append(Vars.DelFileList, fileName)
	} else {
		fmt.Println("文件" + fileName + "内容为空或者处理失败")
	}
	wg.Done()
	<-ch
}

func init() {
	// 已 只写入文件|没有时创建|文件尾部追加 的形式打开这个文件
	flag := os.O_WRONLY | os.O_CREATE | os.O_APPEND
	if logFile, err := os.OpenFile(`./program.log`, flag, 0666); err == nil {
		log.SetOutput(logFile)
	}
	// 设置存储位置、
	Vars = config.InitConfig()
	config.MkdirFile(Vars.DocxFileName)
	config.MkdirFile(Vars.TextFileName)
	if err := json.Unmarshal(Vars.FileInformation, &Vars.FileStruct); err != nil {
		log.Fatalf("error unmarshaling: %s", err)
	}
	if Vars.FileStruct.DelDocxFile {
		fmt.Println("[提醒]已经开启转换后删除旧docx文件选择")
	}
	if Vars.FileStruct.DelDocFile {
		fmt.Println("[提醒]已经开启转换后删除旧doc文件选择")
	}
	if Vars.FileStruct.DocToDocx {
		fmt.Println("[提醒]已经开启doc转换成docx选择")
	}
}

func delDocxFile() {
	for _, fileName := range Vars.DelFileList {
		if err := os.Remove(Vars.DocxFileName + fileName); err != nil {
			log.Println(err)
		}
	}
	fmt.Println("已经删除所有docx文件")
	Vars.DelFileList = []string{} // 清空删除文件列表
}

func main() {
	ch, wg := make(chan struct{}, 3), sync.WaitGroup{}
	if Vars.FileStruct.DocToDocx {
		// 调用python脚本转换doc为docx
		config.CmdPythonSaveDocx([]string{"run.py", Vars.DocxFileName, Vars.TextFileName})
		Vars.FileNameList = config.FileNameList(Vars.DocxFileName) // 重新获取当前目录下所有文件名
	}
	if NameList := Vars.FileNameList; NameList != nil || len(NameList) != 0 {
		for index, file := range Vars.FileNameList {
			fileName := filepath.Base(file)
			switch path.Ext(fileName) {
			case ".docx":
				ch <- struct{}{}
				wg.Add(1)
				go getDocxInformation(fileName, index, ch, &wg)
			default:
				continue
				//fmt.Println("No:", index, fileName, "不是docx文件，跳过处理！")
			}
		}
		wg.Wait()
		fmt.Println("文档转换处理完成！")
		if Vars.FileStruct.DelDocxFile && len(Vars.DelFileList) != 0 {
			fmt.Println("[提醒]开始删除旧docx文件")
			delDocxFile()
		}
	}
	log.Println("文件列表获取失败或没有查找到doc docx 文档")
}
