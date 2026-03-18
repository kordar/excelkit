package excelkit

// --------------------
// SheetBuilder / ColumnBuilder
// --------------------
type SheetBuilder[T any] struct {
	parent       *ExportBuilder[T]
	name         string
	columns      []*ColumnBuilder[T]
	headerStyle  *Style
	defaultStyle *Style
}

func (s *SheetBuilder[T]) Column(header string) *ColumnBuilder[T] {
	c := &ColumnBuilder[T]{parent: s, header: header}
	s.columns = append(s.columns, c)
	return c
}

func (s *SheetBuilder[T]) HeaderStyle(style *Style) *SheetBuilder[T] {
	s.headerStyle = style
	return s
}

func (s *SheetBuilder[T]) SheetDefaultStyle(style *Style) *SheetBuilder[T] {
	s.defaultStyle = style
	return s
}

func (s *SheetBuilder[T]) EndSheet() *ExportBuilder[T] { return s.parent }

func (s *SheetBuilder[T]) ToSchema() SheetSchema[T] {
	cols := make([]Column[T], 0, len(s.columns))
	for _, c := range s.columns {
		cols = append(cols, Column[T]{
			Header: c.header,
			Width:  c.width,
			Value:  c.value,
			Style:  c.style,
			Render: c.render,
		})
	}
	return SheetSchema[T]{
		Name:         s.name,
		Columns:      cols,
		HeaderStyle:  s.headerStyle,
		DefaultStyle: s.defaultStyle,
	}
}

type ColumnBuilder[T any] struct {
	parent *SheetBuilder[T]
	header string
	width  float64
	value  func(T) any
	style  func(T) *Style
	render func(*CellContext[T]) error
}

func (c *ColumnBuilder[T]) Width(w float64) *ColumnBuilder[T]         { c.width = w; return c }
func (c *ColumnBuilder[T]) Value(fn func(T) any) *ColumnBuilder[T]    { c.value = fn; return c }
func (c *ColumnBuilder[T]) Style(fn func(T) *Style) *ColumnBuilder[T] { c.style = fn; return c }
func (c *ColumnBuilder[T]) Render(fn func(*CellContext[T]) error) *ColumnBuilder[T] {
	c.render = fn
	return c
}
func (c *ColumnBuilder[T]) End() *SheetBuilder[T] { return c.parent }
