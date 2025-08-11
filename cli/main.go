//go:build !js || !wasm
// +build !js !wasm

package main

import (
	"github.com/gerblesh/wpi-sched/cmd"
)

func main() {
	cmd.Execute()
}
