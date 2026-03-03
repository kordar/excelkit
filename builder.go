package excelkit

import (
	"io"
	"net/http"
	"reflect"

	"github.com/xuri/excelize/v2"
)

// --------------------
// Builder / DSL
// --------------------
type ExportBuilder[T any] struct {
	File         *excelize.File
	source       DataSource[T]
	sheets       []*SheetBuilder[T]
	useStream    bool
	interceptors []RowInterceptor[T]
}

func New[T any]() *ExportBuilder[T] {
	return &ExportBuilder[T]{File: excelize.NewFile()}
}

func (b *ExportBuilder[T]) UseStream() *ExportBuilder[T] { b.useStream = true; return b }
func (b *ExportBuilder[T]) FromSlice(data []T) *ExportBuilder[T] {
	b.source = &SliceSource[T]{data: data}
	return b
}

func (b *ExportBuilder[T]) Sheet(name string) *SheetBuilder[T] {
	sheet := &SheetBuilder[T]{parent: b, name: name}
	b.sheets = append(b.sheets, sheet)
	return sheet
}

// AutoSheet 自动生成列
func (b *ExportBuilder[T]) AutoSheet(tag string, name string) *SheetBuilder[T] {
	s := b.Sheet(name)
	var t T
	typ := reflect.TypeOf(t)
	if typ.Kind() == reflect.Ptr {
		typ = typ.Elem()
	}
	for i := 0; i < typ.NumField(); i++ {
		f := typ.Field(i)
		colName := f.Name
		if v, ok := f.Tag.Lookup(tag); ok {
			colName = v
		}
		idx := i
		s.Column(colName).Value(func(item T) any {
			val := reflect.ValueOf(item)
			if val.Kind() == reflect.Ptr {
				val = val.Elem()
			}
			return val.Field(idx).Interface()
		}).End()
	}
	return s
}

func (b *ExportBuilder[T]) AddInterceptor(fn RowInterceptor[T]) *ExportBuilder[T] {
	b.interceptors = append(b.interceptors, fn)
	return b
}

func (b *ExportBuilder[T]) build() error {
	for _, sheet := range b.sheets {
		writer := ExcelWriter[T]{
			File:         b.File,
			Schema:       sheet.ToSchema(),
			Source:       b.source,
			UseStream:    b.useStream,
			Interceptors: b.interceptors,
		}
		if err := writer.Write(); err != nil {
			return err
		}
	}
	return nil
}

func (b *ExportBuilder[T]) Save(filename string) error {
	if err := b.build(); err != nil {
		return err
	}
	return b.File.SaveAs(filename)
}

func (b *ExportBuilder[T]) Write(w io.Writer) error {
	if err := b.build(); err != nil {
		return err
	}
	return b.File.Write(w)
}

func (b *ExportBuilder[T]) Download(w http.ResponseWriter, filename string, headers ...http.Header) error {
	w.Header().Set("Content-Type", "application/octet-stream")
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	w.Header().Set("Content-Transfer-Encoding", "binary")
	w.Header().Set("Access-Control-Expose-Headers", "Content-Disposition")

	for _, h := range headers {
		for k, v := range h {
			w.Header()[k] = v
		}
	}

	return b.Write(w)
}
