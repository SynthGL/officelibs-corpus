// officecorpus one-edit round trip for Excelize (github.com/xuri/excelize/v2).
//
// Serve protocol (one long-lived process per container, many files): each stdin line is
// "<ext>\t<base64 input>\t<base64 edit JSON>". The helper writes the input to a private file,
// opens it with excelize.OpenFile(path) using default options, applies the set-cell edit with
// f.SetCellStr(sheet, cell, value), writes it with f.SaveAs(output), closes it, and answers with
// JSON lines: opened, saved or error (phase open, edit or save), then done carrying elapsed_ms (open,
// edit and save, in process) and the base64 output.
package main

import (
	"bufio"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"runtime/debug"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

var out = bufio.NewWriter(os.Stdout)

// editSpec is the runner's precomputed one-edit spec for an xlsx input.
type editSpec struct {
	Format string `json:"format"`
	Kind   string `json:"kind"`
	Sheet  string `json:"sheet"`
	Cell   string `json:"cell"`
	Value  string `json:"value"`
}

func emit(payload map[string]any) {
	data, _ := json.Marshal(payload)
	out.Write(data)
	out.WriteByte('\n')
	out.Flush()
}

func fail(phase string, err error, trace string) {
	message := err.Error()
	if len(message) > 2000 {
		message = message[:2000]
	}
	if len(trace) > 4000 {
		trace = trace[len(trace)-4000:]
	}
	emit(map[string]any{"event": "error", "phase": phase, "type": fmt.Sprintf("%T", err), "message": message, "traceback": trace})
}

func version() string {
	info, ok := debug.ReadBuildInfo()
	if ok {
		for _, dep := range info.Deps {
			if dep.Path == "github.com/xuri/excelize/v2" {
				return strings.TrimPrefix(dep.Version, "v")
			}
		}
	}
	return "unknown"
}

// guard converts a library panic into an error event for the phase.
func guard(phase string, ok *bool) {
	if r := recover(); r != nil {
		fail(phase, fmt.Errorf("panic: %v", r), string(debug.Stack()))
		*ok = false
	}
}

func open(source string) (f *excelize.File, ok bool) {
	ok = true
	defer guard("open", &ok)
	f, err := excelize.OpenFile(source)
	if err != nil {
		fail("open", err, "")
		return nil, false
	}
	return f, ok
}

func edit(f *excelize.File, raw []byte) (ok bool) {
	ok = true
	defer guard("edit", &ok)
	var spec editSpec
	if err := json.Unmarshal(raw, &spec); err != nil {
		fail("edit", err, "")
		return false
	}
	if spec.Format != "xlsx" || spec.Kind != "set-cell" {
		fail("edit", fmt.Errorf("unsupported edit %s/%s", spec.Format, spec.Kind), "")
		return false
	}
	if err := f.SetCellStr(spec.Sheet, spec.Cell, spec.Value); err != nil {
		fail("edit", err, "")
		return false
	}
	return ok
}

func save(f *excelize.File, target string) (ok bool) {
	ok = true
	defer guard("save", &ok)
	if err := f.SaveAs(target); err != nil {
		fail("save", err, "")
		return false
	}
	if err := f.Close(); err != nil {
		fail("save", err, "")
		return false
	}
	return ok
}

func roundTrip(ext, source, target string, spec []byte) {
	if ext != "xlsx" {
		fail("open", fmt.Errorf("unsupported format %s", ext), "")
		return
	}
	f, ok := open(source)
	if !ok {
		return
	}
	emit(map[string]any{"event": "opened"})
	if spec == nil {
		fail("edit", fmt.Errorf("request carries no edit spec"), "")
		f.Close()
		return
	}
	if !edit(f, spec) {
		f.Close()
		return
	}
	if save(f, target) {
		emit(map[string]any{"event": "saved"})
	}
}

func main() {
	if len(os.Args) == 2 && os.Args[1] == "--version" {
		emit(map[string]any{"event": "version", "version": version()})
		return
	}
	work, err := os.MkdirTemp("", "officecorpus")
	if err != nil {
		panic(err)
	}
	in := bufio.NewReaderSize(os.Stdin, 1<<20)
	for {
		line, err := in.ReadString('\n')
		line = strings.TrimRight(line, "\n")
		if line != "" {
			fields := strings.Split(line, "\t")
			ext := fields[0]
			data, decodeErr := base64.StdEncoding.DecodeString(fields[1])
			if decodeErr != nil {
				panic(decodeErr)
			}
			var spec []byte
			if len(fields) > 2 {
				if spec, decodeErr = base64.StdEncoding.DecodeString(fields[2]); decodeErr != nil {
					panic(decodeErr)
				}
			}
			source := filepath.Join(work, "input."+ext)
			target := filepath.Join(work, "output."+ext)
			os.Remove(target)
			if writeErr := os.WriteFile(source, data, 0o600); writeErr != nil {
				panic(writeErr)
			}
			start := time.Now()
			roundTrip(ext, source, target, spec)
			elapsed := float64(time.Since(start).Microseconds()) / 1000
			output := ""
			if produced, readErr := os.ReadFile(target); readErr == nil {
				output = base64.StdEncoding.EncodeToString(produced)
			}
			emit(map[string]any{"event": "done", "elapsed_ms": elapsed, "output": output})
		}
		if err != nil {
			return
		}
	}
}
