const std = @import("std");
const xlsxio = @import("xlsxio");
const odbc = @import("odbc");

pub const SheetData = struct {
    arena: std.heap.ArenaAllocator,
    rows: []const []const []const u8,

    pub fn deinit(self: *SheetData) void {
        self.arena.deinit();
    }

    pub fn shape(self: *const SheetData) struct { rows: usize, cols: usize } {
        var cols: usize = 0;
        for (self.rows) |row| {
            if (row.len > cols) cols = row.len;
        }
        return .{ .rows = self.rows.len, .cols = cols };
    }

    pub fn columns(self: *const SheetData) []const []const u8 {
        if (self.rows.len == 0) return &[_][]const u8{};
        return self.rows[0];
    }

    pub fn locRow(self: *const SheetData, row_index: usize) ?[]const []const u8 {
        if (row_index >= self.rows.len) return null;
        return self.rows[row_index];
    }

    pub fn locCell(self: *const SheetData, row_index: usize, col_index: usize) ?[]const u8 {
        if (row_index >= self.rows.len) return null;
        const row = self.rows[row_index];
        if (col_index >= row.len) return null;
        return row[col_index];
    }

    pub fn locCellByName(self: *const SheetData, row_index: usize, col_name: []const u8) ?[]const u8 {
        if (self.rows.len == 0) return null;
        const headers = self.rows[0];
        for (headers, 0..) |h, c| {
            if (std.mem.eql(u8, h, col_name)) {
                if (row_index >= self.rows.len) return null;
                const row = self.rows[row_index];
                if (c >= row.len) return null;
                return row[c];
            }
        }
        return null;
    }

    pub fn to_excel(self: *const SheetData, allocator: std.mem.Allocator, filename: []const u8, sheet_name: []const u8) !void {
        const filename_z = try allocator.dupeZ(u8, filename);
        defer allocator.free(filename_z);
        const sheet_name_z = try allocator.dupeZ(u8, sheet_name);
        defer allocator.free(sheet_name_z);

        var writer = try xlsxio.Writer.init(filename_z, sheet_name_z);
        defer writer.deinit();

        var buf = std.ArrayList(u8).empty;
        defer buf.deinit(allocator);

        for (self.rows) |row| {
            for (row) |cell| {
                buf.clearRetainingCapacity();
                try buf.appendSlice(allocator, cell);
                try buf.append(allocator, 0);
                const cell_z: [:0]const u8 = buf.items[0..cell.len :0];
                writer.addCellString(cell_z);
            }
            writer.nextRow();
        }
    }

    pub fn to_csv(self: *const SheetData, allocator: std.mem.Allocator, filename: []const u8) !void {
        var buf = std.ArrayList(u8).empty;
        defer buf.deinit(allocator);

        for (self.rows) |row| {
            var col_i: usize = 0;
            for (row) |cell| {
                if (col_i > 0) try buf.append(allocator, ',');
                col_i += 1;
                const needs_quotes = std.mem.indexOfAny(u8, cell, ",\"\r\n") != null;
                if (needs_quotes) {
                    try buf.append(allocator, '"');
                    for (cell) |b| {
                        if (b == '"')
                            try buf.appendSlice(allocator, "\"\"")
                        else
                            try buf.append(allocator, b);
                    }
                    try buf.append(allocator, '"');
                } else {
                    try buf.appendSlice(allocator, cell);
                }
            }
            try buf.append(allocator, '\n');
        }

        const filename_z = try allocator.dupeZ(u8, filename);
        defer allocator.free(filename_z);
        const f = std.c.fopen(filename_z, "wb") orelse return error.FileNotFound;
        defer _ = std.c.fclose(f);
        const slice = buf.items;
        if (slice.len > 0 and std.c.fwrite(slice.ptr, 1, slice.len, f) != slice.len)
            return error.WriteError;
    }

    pub fn head(self: *const SheetData, allocator: std.mem.Allocator, n: ?usize) void {
        const n_rows = n orelse 5;
        const rows_to_show = self.rows[0..@min(n_rows, self.rows.len)];
        if (rows_to_show.len == 0) return;

        var max_cols: usize = 0;
        for (rows_to_show) |row| {
            if (row.len > max_cols) max_cols = row.len;
        }
        if (max_cols == 0) return;

        var col_widths: [512]usize = [_]usize{0} ** 512;
        const col_limit = @min(max_cols, col_widths.len);
        for (rows_to_show) |row| {
            for (row, 0..) |cell, c| {
                if (c >= col_limit) break;
                const w = std.unicode.utf8CountCodepoints(cell) catch cell.len;
                if (w > col_widths[c]) col_widths[c] = w;
            }
        }
        for (col_widths[0..col_limit]) |*w| {
            if (w.* < 1) w.* = 1;
        }

        var buf = std.ArrayList(u8).empty;
        defer buf.deinit(allocator);

        const pad = struct {
            fn addCell(list: *std.ArrayList(u8), a: std.mem.Allocator, cell: []const u8, width: usize) !void {
                const len = std.unicode.utf8CountCodepoints(cell) catch cell.len;
                try list.append(a, ' ');
                try list.appendSlice(a, cell);
                var i: usize = len;
                while (i <= width) : (i += 1) try list.append(a, ' ');
            }
        }.addCell;

        buf.append(allocator, '+') catch return;
        for (col_widths[0..col_limit]) |col_w| {
            buf.append(allocator, '-') catch return;
            var i: usize = 0;
            while (i < col_w + 2) : (i += 1) buf.append(allocator, '-') catch return;
        }
        buf.appendSlice(allocator, "+\n") catch return;

        for (rows_to_show) |row| {
            buf.append(allocator, '|') catch return;
            for (col_widths[0..col_limit], 0..) |col_w, c| {
                const cell = if (c < row.len) row[c] else "";
                pad(&buf, allocator, cell, col_w) catch return;
                buf.append(allocator, '|') catch return;
            }
            buf.appendSlice(allocator, "\n") catch return;
        }

        buf.append(allocator, '+') catch return;
        for (col_widths[0..col_limit]) |col_w| {
            buf.append(allocator, '-') catch return;
            var i: usize = 0;
            while (i < col_w + 2) : (i += 1) buf.append(allocator, '-') catch return;
        }
        buf.appendSlice(allocator, "+\n") catch return;

        std.debug.print("{s}", .{buf.items});
    }

    /// Print the last n rows to the terminal in a table. Default n = 5. Uses allocator for a temporary buffer.
    pub fn tail(self: *const SheetData, allocator: std.mem.Allocator, n: ?usize) void {
        const n_rows = n orelse 5;
        const start = if (self.rows.len <= n_rows) 0 else self.rows.len - n_rows;
        const rows_to_show = self.rows[start..];
        if (rows_to_show.len == 0) return;

        var max_cols: usize = 0;
        for (rows_to_show) |row| {
            if (row.len > max_cols) max_cols = row.len;
        }
        if (max_cols == 0) return;

        var col_widths: [512]usize = [_]usize{0} ** 512;
        const col_limit = @min(max_cols, col_widths.len);
        for (rows_to_show) |row| {
            for (row, 0..) |cell, c| {
                if (c >= col_limit) break;
                const w = std.unicode.utf8CountCodepoints(cell) catch cell.len;
                if (w > col_widths[c]) col_widths[c] = w;
            }
        }
        for (col_widths[0..col_limit]) |*w| {
            if (w.* < 1) w.* = 1;
        }

        var buf = std.ArrayList(u8).empty;
        defer buf.deinit(allocator);

        const pad = struct {
            fn addCell(list: *std.ArrayList(u8), a: std.mem.Allocator, cell: []const u8, width: usize) !void {
                const len = std.unicode.utf8CountCodepoints(cell) catch cell.len;
                try list.append(a, ' ');
                try list.appendSlice(a, cell);
                var i: usize = len;
                while (i <= width) : (i += 1) try list.append(a, ' ');
            }
        }.addCell;

        buf.append(allocator, '+') catch return;
        for (col_widths[0..col_limit]) |col_w| {
            buf.append(allocator, '-') catch return;
            var i: usize = 0;
            while (i < col_w + 2) : (i += 1) buf.append(allocator, '-') catch return;
        }
        buf.appendSlice(allocator, "+\n") catch return;

        for (rows_to_show) |row| {
            buf.append(allocator, '|') catch return;
            for (col_widths[0..col_limit], 0..) |col_w, c| {
                const cell = if (c < row.len) row[c] else "";
                pad(&buf, allocator, cell, col_w) catch return;
                buf.append(allocator, '|') catch return;
            }
            buf.appendSlice(allocator, "\n") catch return;
        }

        buf.append(allocator, '+') catch return;
        for (col_widths[0..col_limit]) |col_w| {
            buf.append(allocator, '-') catch return;
            var i: usize = 0;
            while (i < col_w + 2) : (i += 1) buf.append(allocator, '-') catch return;
        }
        buf.appendSlice(allocator, "+\n") catch return;

        std.debug.print("{s}", .{buf.items});
    }

    /// Print a summary of the sheet (row count, columns, non-null counts, dtypes, memory). Uses allocator for a temporary buffer.
    pub fn info(self: *const SheetData, allocator: std.mem.Allocator) void {
        const n_rows = self.rows.len;
        var max_cols: usize = 0;
        for (self.rows) |row| {
            if (row.len > max_cols) max_cols = row.len;
        }
        if (max_cols > 512) return;
        const n_cols = max_cols;

        var non_null: [512]usize = [_]usize{0} ** 512;
        var total_bytes: usize = 0;
        for (self.rows) |row| {
            for (row, 0..) |cell, c| {
                if (c >= 512) break;
                if (cell.len > 0) non_null[c] += 1;
                total_bytes += cell.len;
            }
        }

        const Dtype = enum { int64, float64, object };
        var dtypes: [512]Dtype = [_]Dtype{.object} ** 512;
        for (0..n_cols) |c| {
            var all_int = true;
            var any_float = false;
            for (self.rows) |row| {
                if (c >= row.len) continue;
                const cell = row[c];
                if (cell.len == 0) continue;
                if (std.fmt.parseInt(i64, cell, 10)) |_| {} else |_| {
                    if (std.fmt.parseFloat(f64, cell)) |_| {
                        any_float = true;
                        all_int = false;
                    } else |_| {
                        all_int = false;
                        any_float = false;
                        break;
                    }
                }
            }
            if (any_float) dtypes[c] = .float64 else if (all_int) dtypes[c] = .int64 else dtypes[c] = .object;
        }

        var buf = std.ArrayList(u8).empty;
        defer buf.deinit(allocator);

        const line = std.fmt.allocPrint(allocator, "SheetData summary\n  Rows: {d}\n  Columns: {d}\n  Memory: ~{d} bytes ({d:.1} KiB)\n\n", .{
            n_rows,
            n_cols,
            total_bytes,
            @as(f64, @floatFromInt(total_bytes)) / 1024.0,
        }) catch return;
        defer allocator.free(line);
        buf.appendSlice(allocator, line) catch return;

        buf.appendSlice(allocator, "  Column              non-null  dtype\n") catch return;
        buf.appendSlice(allocator, "  ------------------  --------  ------\n") catch return;

        const headers = if (n_rows > 0) self.rows[0] else &[_][]const u8{};
        var tmp: [128]u8 = undefined;
        for (0..n_cols) |c| {
            const name: []const u8 = if (c < headers.len and headers[c].len > 0)
                headers[c]
            else
                std.fmt.bufPrint(&tmp, "#{d}", .{c}) catch return;
            const name_pad = if (name.len > 20) name[0..20] else name;
            var name_line: [32]u8 = undefined;
            @memset(name_line[0..], ' ');
            const n_copy = @min(name_pad.len, 20);
            for (name_pad[0..n_copy], name_line[0..n_copy]) |b, *out| out.* = b;
            const nn = non_null[c];
            const dt: []const u8 = switch (dtypes[c]) {
                .int64 => "int64",
                .float64 => "float64",
                .object => "object",
            };
            const row_line = std.fmt.bufPrint(&tmp, "  {s}  {d:>8}  {s}\n", .{ name_line[0..20], nn, dt }) catch return;
            buf.appendSlice(allocator, row_line) catch return;
        }
        buf.append(allocator, '\n') catch return;

        std.debug.print("{s}", .{buf.items});
    }
};

pub fn read_excel(backing_allocator: std.mem.Allocator, filename: []const u8, sheet_name: []const u8) !SheetData {
    var arena = std.heap.ArenaAllocator.init(backing_allocator);
    errdefer arena.deinit();

    const a = arena.allocator();

    const filename_z = try a.dupeZ(u8, filename);
    const sheet_name_z = try a.dupeZ(u8, sheet_name);

    var reader = try xlsxio.Reader.init(a, filename_z);
    defer reader.deinit();

    var sheet = try xlsxio.Reader.Sheet.init(&reader, sheet_name_z);
    defer sheet.deinit();

    var row_list = std.ArrayList([]const []const u8).empty;

    while (sheet.nextRow()) {
        var col_list = std.ArrayList([]const u8).empty;
        while (true) {
            const cell = try sheet.nextCellString();
            if (cell == null) break;
            const cell_value = cell.?;
            defer a.free(cell_value);
            const slice: []const u8 = cell_value[0..cell_value.len];
            const dup = try a.dupe(u8, slice);
            try col_list.append(a, dup);
        }
        try row_list.append(a, try col_list.toOwnedSlice(a));
    }

    return .{
        .arena = arena,
        .rows = try row_list.toOwnedSlice(a),
    };
}

pub fn read_sql(backing_allocator: std.mem.Allocator, conn: *odbc.Connection, query: []const u8) !SheetData {
    var arena = std.heap.ArenaAllocator.init(backing_allocator);
    errdefer arena.deinit();
    const a = arena.allocator();

    var result = conn.query(query) catch |err| return err;
    defer result.deinit();

    var row_list = std.ArrayList([]const []const u8).empty;

    // Header row: column names from the result set
    var header_list = std.ArrayList([]const u8).empty;
    for (1..result.columnCount() + 1) |c| {
        const name = result.columnName(a, @intCast(c)) catch (try a.dupe(u8, ""));
        try header_list.append(a, name);
    }
    try row_list.append(a, try header_list.toOwnedSlice(a));

    while (result.next()) |*row| {
        var col_list = std.ArrayList([]const u8).empty;
        for (1..result.columnCount() + 1) |c| {
            const val = row.get(@intCast(c)) catch null;
            defer if (val) |v| conn.allocator.free(v);
            const cell_slice = val orelse "";
            const dup = try a.dupe(u8, cell_slice);
            try col_list.append(a, dup);
        }
        try row_list.append(a, try col_list.toOwnedSlice(a));
    }

    return .{
        .arena = arena,
        .rows = try row_list.toOwnedSlice(a),
    };
}

pub fn read_csv(backing_allocator: std.mem.Allocator, filename: []const u8) !SheetData {
    var arena = std.heap.ArenaAllocator.init(backing_allocator);
    errdefer arena.deinit();
    const a = arena.allocator();

    const filename_z = try a.dupeZ(u8, filename);
    const f = std.c.fopen(filename_z, "rb") orelse return error.FileNotFound;
    defer _ = std.c.fclose(f);

    var content = std.ArrayList(u8).empty;
    defer content.deinit(a);
    var buf: [4096]u8 = undefined;
    while (true) {
        const n = std.c.fread(buf[0..].ptr, 1, buf.len, f);
        if (n == 0) break;
        try content.appendSlice(a, buf[0..n]);
        if (n < buf.len) break;
    }

    var row_list = std.ArrayList([]const []const u8).empty;
    var i: usize = 0;
    const data = content.items;

    while (i < data.len) {
        var col_list = std.ArrayList([]const u8).empty;
        while (i < data.len) {
            const cell_start = i;
            if (data[i] == '"') {
                i += 1;
                var cell_buf = std.ArrayList(u8).empty;
                defer cell_buf.deinit(a);
                while (i < data.len) {
                    if (data[i] == '"') {
                        i += 1;
                        if (i < data.len and data[i] == '"') {
                            try cell_buf.append(a, '"');
                            i += 1;
                        } else break;
                    } else {
                        try cell_buf.append(a, data[i]);
                        i += 1;
                    }
                }
                try col_list.append(a, try a.dupe(u8, cell_buf.items));
            } else {
                while (i < data.len and data[i] != ',' and data[i] != '\n' and data[i] != '\r') i += 1;
                try col_list.append(a, try a.dupe(u8, data[cell_start..i]));
            }
            if (i < data.len and data[i] == ',') {
                i += 1;
            } else {
                if (i < data.len and (data[i] == '\r' or data[i] == '\n')) {
                    i += 1;
                    if (i < data.len and data[i] == '\n') i += 1;
                }
                break;
            }
        }
        if (col_list.items.len > 0)
            try row_list.append(a, try col_list.toOwnedSlice(a));
    }

    return .{
        .arena = arena,
        .rows = try row_list.toOwnedSlice(a),
    };
}
