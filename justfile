default:
    just --list

build-cli:
    go build -o "wpi-sched" ./cli

build-wasm:
    #!/usr/bin/env bash
    GOOS=js GOARCH=wasm go build -o static/main.wasm ./wasm
    cp "$(go env GOROOT)/lib/wasm/wasm_exec.js" static/

python-http:
    python -m http.server
