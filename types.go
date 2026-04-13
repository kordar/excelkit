package excelkit

import "github.com/xuri/excelize/v2"

// --------------------
// CellContext
// --------------------
type CellContext[T any] struct {
	File     *excelize.File
	Sheet    string
	RowIndex int
	ColIndex int
	RowData  T
}

// --------------------
// Column / SheetSchema
// --------------------
type Column[T any] struct {
	Header string
	Width  float64
	Value  func(T) any
	Style  func(T) *Style
	Render func(*CellContext[T]) error
}

type SheetSchema[T any] struct {
	Name          string
	Columns       []Column[T]
	Title         string
	Subtitle      string
	TitleStyle    *Style
	SubtitleStyle *Style
	HeaderStyle   *Style
	DefaultStyle  *Style
}

// --------------------
// 数据源接口
// --------------------
type DataSource[T any] interface {
	Next() (T, bool, error)
	Close() error
}

type SliceSource[T any] struct {
	data []T
	idx  int
}

func (s *SliceSource[T]) Next() (T, bool, error) {
	if s.idx >= len(s.data) {
		var zero T
		return zero, false, nil
	}
	v := s.data[s.idx]
	s.idx++
	return v, true, nil
}

func (s *SliceSource[T]) Close() error { return nil }

// --------------------
// 高级行拦截器
// --------------------
type MergeRange[T any] struct {
	StartRow int
	EndRow   int
	StartCol int
	EndCol   int
	RowData  T
	Style    *Style
}

type RowInterceptor[T any] func(row T, rowIndex int) (write bool, extraRows []T, merge []MergeRange[T])
