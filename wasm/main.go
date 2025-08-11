//go:build js && wasm
// +build js,wasm

package main

import (
	"bytes"
	"syscall/js"

	"github.com/gerblesh/wpi-sched/cmd"
	"github.com/xuri/excelize/v2"
)

func main() {
	done := make(chan struct{})
	js.Global().Set("processFile", js.FuncOf(processFile))
	<-done
}

func processFile(this js.Value, args []js.Value) any {
	if len(args) < 1 {
		return js.ValueOf("missing file data")
	}

	fileBytes := make([]byte, args[0].Length())
	js.CopyBytesToGo(fileBytes, args[0])
	r := bytes.NewReader(fileBytes)
	f, err := excelize.OpenReader(r)
	if err != nil {
		return js.ValueOf("error: " + err.Error())
	}
	courses, err := cmd.GetCourses(f)
	if err != nil {
		return js.ValueOf("error: " + err.Error())
	}

	var buf bytes.Buffer
	err = cmd.WriteIcalBuf(courses, &buf)
	if err != nil {
		return js.ValueOf("error: " + err.Error())
	}

	outJS := js.Global().Get("Uint8Array").New(len(buf.Bytes()))
	js.CopyBytesToJS(outJS, buf.Bytes())
	return outJS
}
