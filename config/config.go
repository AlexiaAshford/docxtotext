package config

import (
	"encoding/json"
	"io/ioutil"
	"log"
	"os"
)

type Config struct {
	DocToDocx   bool
	DelDocFile  bool
	DelDocxFile bool
}

type ClassConfig struct {
	ConfigFile      string
	FileInformation []byte
	FileNameList    []string
	FileStruct      Config
	DelFileList     []string
}

func (is *ClassConfig) SaveConfig() {
	if err := ioutil.WriteFile("./config.json", is.FileInformation, 0777); err != nil {
		log.Fatalf("error writing file: %s", err)
	}
}
func (is *ClassConfig) load() {
	if data, err := ioutil.ReadFile("./config.json"); err != nil {
		is.FileInformation = data
	}
}

func InitConfig() *ClassConfig {
	Vars := ClassConfig{ConfigFile: "./config.json"}
	if _, err := os.Stat(Vars.ConfigFile); err != nil {
		if configs, ok := json.MarshalIndent(&Config{}, "", "   "); ok == nil {
			Vars.FileInformation = configs
			Vars.SaveConfig()
		} else {
			log.Fatalf("error marshal config: %s", ok)
		}
	} else {
		if data, err := ioutil.ReadFile("./config.json"); err == nil {
			Vars.FileInformation = data
		} else {
			log.Fatalf("error reading file: %s", err)
		}
	}
	Vars.FileNameList = FileNameList() // 获取当前目录下所有文件名
	return &Vars
}
