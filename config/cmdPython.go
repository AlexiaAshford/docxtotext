package config

import (
	"fmt"
	"log"
	"os/exec"
)

func CmdPythonSaveDocx(arg []string) {
	if cmd, err := exec.Command("python", arg...).Output(); err == nil {
		fmt.Println("doc转换成功：")
		log.Println(string(cmd))
	} else {
		log.Fatal("doc转换失败")
	}

}
