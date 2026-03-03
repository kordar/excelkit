# ExcelKit

ExcelKit 是一个基于 [excelize](https://github.com/xuri/excelize) 的 Go 语言 Excel 导出工具库。它提供了流畅的链式调用 API（DSL），支持结构体自动映射、流式写入以及强大的行拦截器功能（支持插入小结行、合并单元格等）。

## 特性

- **链式调用（DSL）**：简洁优雅的 API 设计，易于阅读和维护。
- **泛型支持**：基于 Go 1.18+ 泛型，类型安全。
- **流式写入**：支持 `UseStream()` 模式，能够高效处理大数据量导出。
- **自动映射**：支持通过 Struct Tag 自动生成表头和数据列。
- **多 Sheet 支持**：可以在一个文件中创建多个不同的数据表。
- **高度定制**：
    - 自定义列值转换。
    - 自定义样式（字体、背景、边框等）。
    - 强大的拦截器（Interceptor）：可在任意位置插入额外行（如小结、合计），并支持单元格合并。

## 安装

```bash
go get github.com/kordar/excelkit
```

## 快速开始

### 1. 定义模型

```go
type User struct {
    ID        int64     `json:"id" excel:"ID"`
    Name      string    `json:"name" excel:"姓名"`
    Email     string    `json:"email" excel:"邮箱"`
    Status    string    `json:"status" excel:"状态"`
    CreatedAt time.Time `json:"created_at" excel:"创建时间"`
}
```

### 2. 简单导出（自动映射）

```go
func main() {
    users := []User{...} // 数据源

    err := excelkit.New[User]().
        FromSlice(users).
        AutoSheet("excel", "用户表"). // 使用 `excel` tag 自动生成列
        EndSheet().
        Save("users.xlsx")

    if err != nil {
        panic(err)
    }
}
```

### 3. 高级用法（自定义列与拦截器）

以下示例展示了如何：
- 手动定义列并自定义值逻辑。
- 使用流式写入 (`UseStream`)。
- 使用拦截器每 10 行插入一行“小结”，并合并单元格。

```go
func main() {
    // 1. 准备数据
    users := []User{...}

    // 2. 定义累计变量（用于小结）
    var processedCount int

    // 3. 构建导出器
    err := excelkit.New[User]().
        FromSlice(users).
        UseStream(). // 开启流式写入
        Sheet("用户表").
            Column("ID").Value(func(u User) any {
                if u.IsSummary { return "小结" } // 特殊处理小结行
                return u.ID
            }).End().
            Column("姓名").Value(func(u User) any {
                if u.IsSummary { return u.SummaryCount }
                return u.Name
            }).End().
            Column("邮箱").Value(func(u User) any { return u.Email }).End().
            Column("状态").Value(func(u User) any { return u.Status }).End().
            Column("创建时间").Value(func(u User) any { return u.CreatedAt }).End().
        EndSheet().
        // 添加拦截器
        AddInterceptor(func(row User, rowIndex int) (bool, []User, []MergeRange[User]) {
            processedCount++
            extraRows := []User{}
            mergeRanges := []MergeRange[User]{}

            // 每 10 行插入小结
            if processedCount > 1 && (processedCount-1)%10 == 0 {
                summary := User{
                    IsSummary:    true,
                    SummaryCount: processedCount - 1,
                }
                extraRows = append(extraRows, summary)

                // 合并单元格：从第 2 列到第 5 列
                mergeRanges = append(mergeRanges, MergeRange[User]{
                    StartRow: rowIndex,
                    EndRow:   rowIndex,
                    StartCol: 2,
                    EndCol:   5,
                    RowData:  summary,
                    Style:    excelkit.BlueSummaryStyle(), // 使用预定义样式
                })
            }
            return true, extraRows, mergeRanges
        }).
        Save("users_advanced.xlsx")

    if err != nil {
        panic(err)
    }
}
```

#### 连续小结（两行，列合并）

下面展示每 10 条数据插入两行小结，并分别将两行小结在“姓名至创建时间”列（B..E）进行横向合并；第一行小结使用蓝底白字样式，第二行使用绿色样式：

```go
var processedCount int

err := excelkit.New[User]().
    FromSlice(users).
    UseStream().
    Sheet("用户表").
        Column("ID").Value(func(u User) any {
            if u.IsSummary { return "小结" }
            return u.ID
        }).End().
        Column("姓名").Value(func(u User) any {
            if u.IsSummary { return u.SummaryCount }
            return u.Name
        }).End().
        Column("邮箱").Value(func(u User) any { return u.Email }).End().
        Column("状态").Value(func(u User) any { return u.Status }).End().
        Column("创建时间").Value(func(u User) any { return u.CreatedAt }).End().
    EndSheet().
    AddInterceptor(func(row User, rowIndex int) (bool, []User, []excelkit.MergeRange[User]) {
        processedCount++
        extra := []User{}
        merges := []excelkit.MergeRange[User]{}

        if processedCount > 1 && (processedCount-1)%10 == 0 {
            s1 := User{IsSummary: true, SummaryCount: processedCount - 1}
            s2 := User{IsSummary: true}
            extra = append(extra, s1, s2)

            merges = append(merges,
                excelkit.MergeRange[User]{
                    StartRow: rowIndex, EndRow: rowIndex,
                    StartCol: 2, EndCol: 5,
                    RowData: s1, Style: excelkit.BlueSummaryStyle(),
                },
                excelkit.MergeRange[User]{
                    StartRow: rowIndex + 1, EndRow: rowIndex + 1,
                    StartCol: 2, EndCol: 5,
                    RowData: s2, Style: excelkit.GreenStyle(),
                },
            )
        }
        return true, extra, merges
    }).
    Save("users_advanced_double_summary.xlsx")

if err != nil {
    panic(err)
}
```

### 4. Web 导出（浏览器下载）

ExcelKit 提供了便捷的 `Download` 方法，可直接将 Excel 写入 `http.ResponseWriter` 并自动设置 `Content-Disposition` 等响应头。

```go
func DownloadHandler(w http.ResponseWriter, r *http.Request) {
    users := []User{...}

    // 可选：自定义响应头
    headers := http.Header{}
    headers.Set("X-Custom-Header", "ExcelKit")

    err := excelkit.New[User]().
        FromSlice(users).
        AutoSheet("excel", "用户表").EndSheet().
        Download(w, "users.xlsx", headers) // 自动设置 header 并写入

    if err != nil {
        http.Error(w, err.Error(), http.StatusInternalServerError)
    }
}
```

## 核心 API

- `New[T]()`: 创建一个新的导出构建器。
- `FromSlice([]T)`: 设置切片数据源。
- `UseStream()`: 启用流式写入模式（推荐用于大数据导出）。
- `Sheet(name)`: 开始定义一个 Sheet。
- `AutoSheet(tag, name)`: 根据 Struct Tag 自动生成 Sheet。
- `Column(header)`: 定义一列。
    - `.Value(fn)`: 定义该列取值逻辑。
    - `.Width(w)`: 定义列宽。
    - `.Style(fn)`: 定义动态样式。
- `AddInterceptor(fn)`: 添加行拦截器，用于复杂的行插入和合并逻辑。
- `Save(filename)`: 执行导出并保存文件。
- `Write(io.Writer)`: 将 Excel 内容写入任意 Writer。
- `Download(http.ResponseWriter, filename, ...headers)`: 辅助方法，设置下载响应头并写入 ResponseWriter，支持自定义 Header。

## License

MIT
