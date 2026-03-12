# zig-xlsxio

A Zig wrapper for [xlsxio](https://github.com/brechtsanders/xlsxio), a library for reading and writing Excel XLSX files.

## Features

- Read XLSX files with support for multiple sheets
- Write XLSX files with multiple sheets
- Support for different data types (string, integer, float, datetime)
- Memory-efficient streaming API
- Windows 64-bit support

## Requirements

- Zig 0.14.0 or later
- **Windows 64-bit only** (currently)

## Installation

### Using Zig Package Manager

```bash
# Add to your project
zig fetch --save git+https://github.com/yourusername/zig-xlsxio
```

### Usage in build.zig

The simplest way to use this package is:

```zig
const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    const exe = b.addExecutable(.{
        .name = "my_app",
        .root_source_file = b.path("src/main.zig"),
        .target = target,
        .optimize = optimize,
    });

    // Import xlsxio module
    const xlsxio_dep = b.dependency("xlsxio", .{
        .target = target,
        .optimize = optimize,
    });
    exe.root_module.addImport("xlsxio", xlsxio_dep.module("xlsxio"));
    
    // Link with C libraries (required)
    exe.linkLibC();
    
    // Install the artifact
    b.installArtifact(exe);
    
    // Simple DLL installation - copies DLLs to bin directory
    const bin_dir = xlsxio_dep.path("vendor/xlsxio/bin");
    const dlls = [_][]const u8{
        "xlsxio_read.dll",
        "xlsxio_write.dll",
        "libexpat.dll",
        "minizip.dll", 
        "zlib1.dll",
        "bz2.dll",
    };
    
    for (dlls) |dll| {
        b.installBinFile(b.pathJoin(&.{bin_dir.getPath(b), dll}), dll);
    }
    
    // Create a run command that includes the bin directory in PATH
    const run_cmd = b.addRunArtifact(exe);
    run_cmd.addPathDir("bin");
    
    const run_step = b.step("run", "Run the app");
    run_step.dependOn(&run_cmd.step);
}
```

### Using build_module.zig (Easier Alternative)

For an even easier approach, you can use the provided `build_module.zig` helper:

```zig
const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    const exe = b.addExecutable(.{
        .name = "my_app",
        .root_source_file = b.path("src/main.zig"),
        .target = target,
        .optimize = optimize,
    });

    // Get the xlsxio dependency
    const xlsxio_dep = b.dependency("xlsxio", .{
        .target = target,
        .optimize = optimize,
    });
    
    // Add the build helper module
    const build_mod = @import("xlsxio").build_module;
    
    // One function handles everything: import, linking, and DLL installation
    build_mod.linkXlsxioModule(b, exe, xlsxio_dep);
    
    b.installArtifact(exe);
    
    const run_cmd = b.addRunArtifact(exe);
    const run_step = b.step("run", "Run the app");
    run_step.dependOn(&run_cmd.step);
}
```

### Manual Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/zig-xlsxio.git
cd zig-xlsxio

# Build the project
zig build
```

## Usage

### Reading Excel Files

```zig
const std = @import("std");
const xlsxio = @import("xlsxio");

pub fn main() !void {
    const allocator = std.heap.page_allocator;
    
    // Open Excel file
    var reader = try xlsxio.Reader.init(allocator, "data.xlsx");
    defer reader.deinit();
    
    // Open a sheet (use null for first sheet)
    var sheet = try xlsxio.Reader.Sheet.init(&reader, "Sheet1");
    defer sheet.deinit();
    
    // Iterate through rows
    while (sheet.nextRow()) {
        // Read cells
        const text = try sheet.nextCellString();
        if (text) |t| {
            std.debug.print("Text: {s}\n", .{t});
            allocator.free(t); // Don't forget to free the string!
        }
        
        const number = try sheet.nextCellInt();
        if (number) |n| {
            std.debug.print("Number: {d}\n", .{n});
        }
        
        const float_val = try sheet.nextCellFloat();
        if (float_val) |f| {
            std.debug.print("Float: {d}\n", .{f});
        }
        
        const date = try sheet.nextCellDatetime();
        if (date) |d| {
            std.debug.print("Date (timestamp): {d}\n", .{d.secs});
        }
    }
}
```

### Writing Excel Files

```zig
const std = @import("std");
const xlsxio = @import("xlsxio");

pub fn main() !void {
    // Create a new Excel file
    var writer = try xlsxio.Writer.init("output.xlsx", "Sheet1");
    defer writer.deinit();
    
    // Add column headers
    writer.addCellString("Name");
    writer.addCellString("Age");
    writer.addCellString("Score");
    writer.addCellString("Date");
    writer.nextRow();
    
    // Add data
    writer.addCellString("Alice");
    writer.addCellInt(28);
    writer.addCellFloat(95.5);
    writer.addCellDatetime(xlsxio.Timestamp{ .secs = 1716691200 }); // Unix timestamp
    writer.nextRow();
    
    // Add another sheet
    writer.addSheet("Sheet2");
    
    // Add data to the new sheet
    writer.addCellString("Data on Sheet 2");
    writer.nextRow();
}
```

## API Reference

### Reader

- `Reader.init(allocator, filename)` - Open an Excel file for reading
- `Reader.deinit()` - Close the Excel file
- `Reader.Sheet.init(reader, sheet_name)` - Open a specific sheet (pass null for first sheet)
- `Reader.Sheet.deinit()` - Close the sheet
- `Reader.Sheet.nextRow()` - Move to the next row (returns true if successful)
- `Reader.Sheet.nextCell()` - Get the next cell's content as a raw string
- `Reader.Sheet.nextCellString()` - Get the next cell's content as a string
- `Reader.Sheet.nextCellInt()` - Get the next cell's content as an integer
- `Reader.Sheet.nextCellFloat()` - Get the next cell's content as a float
- `Reader.Sheet.nextCellDatetime()` - Get the next cell's content as a timestamp

### Writer

- `Writer.init(filename, sheet_name)` - Create a new Excel file for writing
- `Writer.deinit()` - Close the Excel file and save changes
- `Writer.addSheet(name)` - Add a new sheet
- `Writer.addCellString(value)` - Add a string cell
- `Writer.addCellInt(value)` - Add an integer cell
- `Writer.addCellFloat(value)` - Add a float cell
- `Writer.addCellDatetime(value)` - Add a datetime cell
- `Writer.nextRow()` - Move to the next row

## Database support (ODBC)

This package includes an **ODBC** module for connecting to databases and running SQL queries. It uses [unixODBC](https://github.com/lurcher/unixODBC) on Linux/macOS and the system ODBC driver manager on Windows.

- **Windows**: links against `odbc32` (built-in). No extra install.
- **Linux/macOS**: install unixODBC first (e.g. `apt install unixodbc-dev` or `brew install unixodbc`), then link with the `odbc` module.

### Usage

```zig
const std = @import("std");
const odbc = @import("odbc");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    var conn = try odbc.Connection.init(allocator);
    defer conn.deinit();

    // Option 1: Connect by DSN
    try conn.connectWithDsn("MyDSN", "user", "password");

    // Option 2: Connect with connection string
    try conn.connectWithConnStr("DRIVER={PostgreSQL};SERVER=localhost;DATABASE=mydb;UID=user;PWD=pass");

    try conn.exec("CREATE TABLE IF NOT EXISTS t (id INT, name VARCHAR(100))");
    try conn.exec("INSERT INTO t VALUES (1, 'hello')");

    var result = try conn.query("SELECT id, name FROM t");
    defer result.deinit();
    while (result.next()) |*row| {
        const id = row.get(1);
        const name = row.get(2);
        if (id) |v| { std.log.info("id: {s}", .{v}); allocator.free(v); }
        if (name) |v| { std.log.info("name: {s}", .{v}); allocator.free(v); }
    }
}
```

Add the `odbc` module in your `build.zig` when using this package as a dependency (it is exported as `odbc`).

### ODBC test executable

Build and run the ODBC integration demo:

```bash
zig build                    # builds odbc_demo into zig-out/bin/
zig build run-odbc -- <connection_string> [query]
```

Example (with a DSN or driver connection string):

```bash
zig build run-odbc -- "DSN=MyDB;UID=user;PWD=pass"
zig build run-odbc -- "DRIVER={ODBC Driver 17 for SQL Server};SERVER=.;DATABASE=test;Trusted_Connection=yes;" "SELECT 1 AS x"
```

## Known Limitations

- Currently only supports Windows 64-bit
- Returned strings must be freed manually

## License

This project is licensed under the MIT License - see the LICENSE file for details.