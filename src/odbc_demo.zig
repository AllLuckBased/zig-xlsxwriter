//! ODBC integration test/demo executable.
//! Fill in server, database, username, password below and run.

const std = @import("std");
const odbc = @import("odbc");
const zigxlsx = @import("zigxlsx");

// --- Fill in your values ---
const server = "server";
const database = "database";
const username = "username";
const password = "password";
const default_sql = "SELECT * FROM PwC_T_EcoResProductColorEntity";
const output_excel = "PwC_T_EcoResProductColorEntity.xlsx";
const output_sheet = "PwC_T_EcoResProductColorEntity";

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    var conn_str_buf: [512]u8 = undefined;
    const conn_str = std.fmt.bufPrint(
        &conn_str_buf,
        "DRIVER={{ODBC Driver 18 for SQL Server}};" ++
            "SERVER={s};DATABASE={s};" ++
            "UID={s};PWD={s};" ++
            "Authentication=ActiveDirectoryPassword;" ++
            "Connection Timeout=120;Login Timeout=120;",
        .{ server, database, username, password },
    ) catch {
        std.log.err("Connection string too long", .{});
        return error.ConnectionStringTooLong;
    };

    std.log.info("Connecting...", .{});
    var conn = odbc.Connection.init(allocator) catch |err| {
        std.log.err("Connection.init failed: {}", .{err});
        return err;
    };
    defer conn.deinit();

    conn.connectWithConnStr(conn_str) catch |err| {
        std.log.err("Connect failed. Check connection string and ODBC driver. Error: {}", .{err});
        return err;
    };
    std.log.info("Connected.", .{});

    std.log.info("Query: {s}", .{default_sql});
    var sheet_data = zigxlsx.read_sql(allocator, &conn, default_sql) catch |err| {
        std.log.err("Query failed: {}", .{err});
        return err;
    };
    defer sheet_data.deinit();

    const sh = sheet_data.shape();
    std.log.info("Rows: {d}, Columns: {d}", .{ sh.rows, sh.cols });

    std.log.info("Saving to {s} (sheet: {s})", .{ output_excel, output_sheet });
    sheet_data.to_excel(allocator, output_excel, output_sheet) catch |err| {
        std.log.err("Failed to save Excel: {}", .{err});
        return err;
    };
    std.log.info("Saved.", .{});
}
