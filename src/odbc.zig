//! Zig bindings for unixODBC / Windows ODBC.
//! Connect to databases and run SQL queries via the standard ODBC API.
//! On Unix/Linux/macOS links against unixODBC (https://github.com/lurcher/unixODBC).
//! On Windows links against the system ODBC driver manager (odbc32).

const std = @import("std");

const c = @cImport({
    @cDefine("ODBCVER", "0x0380");
    @cDefine("SQL_NOUNICODEMAP", "1"); // use ANSI (char*) API so []const u8 works
    @cInclude("sql.h");
    @cInclude("sqlext.h");
});

pub const SQL_SUCCESS = c.SQL_SUCCESS;
pub const SQL_SUCCESS_WITH_INFO = c.SQL_SUCCESS_WITH_INFO;
pub const SQL_ERROR = c.SQL_ERROR;
pub const SQL_NO_DATA = c.SQL_NO_DATA;
pub const SQL_NTS = c.SQL_NTS;
pub const SQL_HANDLE_ENV = c.SQL_HANDLE_ENV;
pub const SQL_HANDLE_DBC = c.SQL_HANDLE_DBC;
pub const SQL_HANDLE_STMT = c.SQL_HANDLE_STMT;
pub const SQL_ATTR_ODBC_VERSION = c.SQL_ATTR_ODBC_VERSION;
pub const SQL_OV_ODBC3 = c.SQL_OV_ODBC3;
pub const SQL_DRIVER_NOPROMPT = c.SQL_DRIVER_NOPROMPT;
pub const SQL_CHAR = c.SQL_CHAR;
pub const SQL_C_CHAR = c.SQL_C_CHAR;
pub const SQL_INTEGER = c.SQL_INTEGER;
pub const SQL_C_SLONG = c.SQL_C_SLONG;

pub const Env = *opaque {};
pub const Dbc = *opaque {};
pub const Stmt = *opaque {};
pub const SQLHENV = c.SQLHENV;
pub const SQLHDBC = c.SQLHDBC;
pub const SQLHSTMT = c.SQLHSTMT;

/// Allocates an environment handle (required before connecting).
pub fn allocEnv() !Env {
    var env: c.SQLHENV = null;
    const ret = c.SQLAllocHandle(c.SQL_HANDLE_ENV, null, &env);
    if (ret != c.SQL_SUCCESS and ret != c.SQL_SUCCESS_WITH_INFO) {
        return error.ODBCAllocEnvFailed;
    }
    return @ptrCast(env);
}

/// Frees an environment handle.
pub fn freeEnv(env: Env) void {
    _ = c.SQLFreeHandle(c.SQL_HANDLE_ENV, @ptrCast(env));
}

/// Allocates a connection handle.
pub fn allocDbc(env: Env) !Dbc {
    var dbc: c.SQLHDBC = null;
    const ret = c.SQLAllocHandle(c.SQL_HANDLE_DBC, @ptrCast(env), &dbc);
    if (ret != c.SQL_SUCCESS and ret != c.SQL_SUCCESS_WITH_INFO) {
        return error.ODBCAllocDbcFailed;
    }
    return @ptrCast(dbc);
}

/// Frees a connection handle.
pub fn freeDbc(dbc: Dbc) void {
    _ = c.SQLFreeHandle(c.SQL_HANDLE_DBC, @ptrCast(dbc));
}

/// Allocates a statement handle (for executing SQL).
pub fn allocStmt(dbc: Dbc) !Stmt {
    var stmt: c.SQLHSTMT = null;
    const ret = c.SQLAllocHandle(c.SQL_HANDLE_STMT, @ptrCast(dbc), &stmt);
    if (ret != c.SQL_SUCCESS and ret != c.SQL_SUCCESS_WITH_INFO) {
        return error.ODBCAllocStmtFailed;
    }
    return @ptrCast(stmt);
}

/// Frees a statement handle.
pub fn freeStmt(stmt: Stmt) void {
    _ = c.SQLFreeHandle(c.SQL_HANDLE_STMT, @ptrCast(stmt));
}

/// Sets the ODBC version on the environment (call before connecting).
pub fn setEnvAttrOdbcVersion(env: Env) !void {
    const ret = c.SQLSetEnvAttr(@ptrCast(env), c.SQL_ATTR_ODBC_VERSION, @ptrFromInt(c.SQL_OV_ODBC3), 0);
    if (ret != c.SQL_SUCCESS and ret != c.SQL_SUCCESS_WITH_INFO) {
        return error.ODBCSetEnvAttrFailed;
    }
}

/// Connects using a DSN (data source name). Pass null for uid/pwd if not needed.
pub fn connectDsn(_: Env, dbc: Dbc, dsn: []const u8, uid: ?[]const u8, pwd: ?[]const u8) !void {
    const dsn_z = try std.heap.page_allocator.dupeZ(u8, dsn);
    defer std.heap.page_allocator.free(dsn_z);
    const uid_z = if (uid) |u| (try std.heap.page_allocator.dupeZ(u8, u)) else (try std.heap.page_allocator.dupeZ(u8, ""));
    defer std.heap.page_allocator.free(uid_z);
    const pwd_z = if (pwd) |p| (try std.heap.page_allocator.dupeZ(u8, p)) else (try std.heap.page_allocator.dupeZ(u8, ""));
    defer std.heap.page_allocator.free(pwd_z);

    const ret = c.SQLConnect(
        @ptrCast(dbc),
        @ptrCast(dsn_z.ptr),
        c.SQL_NTS,
        @ptrCast(uid_z.ptr),
        c.SQL_NTS,
        @ptrCast(pwd_z.ptr),
        c.SQL_NTS,
    );
    if (ret != c.SQL_SUCCESS and ret != c.SQL_SUCCESS_WITH_INFO) {
        return error.ODBCConnectFailed;
    }
}

/// Fetches the first ODBC diagnostic record and prints it to stderr (SQLSTATE, native error, message).
pub fn logLastError(handle_type: c_int, handle: anytype) void {
    var sqlstate: [6]u8 = undefined;
    var native_err: c.SQLINTEGER = 0;
    var msg_buf: [512]u8 = undefined;
    var msg_len: c.SQLSMALLINT = 0;
    const ret = c.SQLGetDiagRec(
        @intCast(handle_type),
        @ptrCast(handle),
        1,
        @ptrCast(&sqlstate),
        &native_err,
        @ptrCast(&msg_buf),
        @intCast(msg_buf.len),
        &msg_len,
    );
    if (ret == c.SQL_SUCCESS or ret == c.SQL_SUCCESS_WITH_INFO) {
        const msg = msg_buf[0..@min(msg_buf.len, @as(usize, @intCast(msg_len)))];
        std.debug.print("ODBC: SQLSTATE={s} NativeError={d} Message={s}\n", .{
            sqlstate[0..5],
            native_err,
            msg,
        });
    }
}

/// Connects using a connection string (e.g. "DRIVER={PostgreSQL};SERVER=localhost;PORT=5432;DATABASE=mydb;UID=user;PWD=pass").
pub fn connectWithString(dbc: Dbc, conn_str: []const u8) !void {
    var out_buf: [1024]u8 = undefined;
    var out_len: c.SQLSMALLINT = 0;
    const conn_z = try std.heap.page_allocator.dupeZ(u8, conn_str);
    defer std.heap.page_allocator.free(conn_z);

    const ret = c.SQLDriverConnect(
        @ptrCast(dbc),
        null,
        @ptrCast(conn_z.ptr),
        c.SQL_NTS,
        @ptrCast(&out_buf),
        @intCast(out_buf.len),
        &out_len,
        c.SQL_DRIVER_NOPROMPT,
    );
    if (ret != c.SQL_SUCCESS and ret != c.SQL_SUCCESS_WITH_INFO) {
        logLastError(c.SQL_HANDLE_DBC, dbc);
        return error.ODBCConnectFailed;
    }
}

/// Disconnects from the database.
pub fn disconnect(dbc: Dbc) void {
    _ = c.SQLDisconnect(@ptrCast(dbc));
}

/// Executes a SQL statement directly (no prepared statement). Use for SELECT/INSERT/UPDATE/DELETE etc.
pub fn execDirect(stmt: Stmt, sql: []const u8) !void {
    const sql_z = try std.heap.page_allocator.dupeZ(u8, sql);
    defer std.heap.page_allocator.free(sql_z);
    const ret = c.SQLExecDirect(@ptrCast(stmt), @ptrCast(sql_z.ptr), c.SQL_NTS);
    if (ret != c.SQL_SUCCESS and ret != c.SQL_SUCCESS_WITH_INFO and ret != c.SQL_NO_DATA) {
        return error.ODBCExecFailed;
    }
}

/// Fetches the next row. Returns true if a row was fetched, false if no more rows.
pub fn fetch(stmt: Stmt) bool {
    const ret = c.SQLFetch(@ptrCast(stmt));
    return ret == c.SQL_SUCCESS or ret == c.SQL_SUCCESS_WITH_INFO;
}

/// Returns the number of columns in the result set.
pub fn numResultCols(stmt: Stmt) !c_int {
    var n: c.SQLSMALLINT = 0;
    const ret = c.SQLNumResultCols(@ptrCast(stmt), &n);
    if (ret != c.SQL_SUCCESS and ret != c.SQL_SUCCESS_WITH_INFO) {
        return error.ODBCNumResultColsFailed;
    }
    return @intCast(n);
}

/// Gets the column name for the given column (1-based). Caller owns the returned slice; use allocator to free.
pub fn getColumnName(allocator: std.mem.Allocator, stmt: Stmt, col: c.SQLUSMALLINT) ![]const u8 {
    var name_buf: [256]u8 = undefined;
    var name_len: c.SQLSMALLINT = 0;
    var data_type: c.SQLSMALLINT = 0;
    var col_size: c.SQLULEN = 0;
    var decimal_digits: c.SQLSMALLINT = 0;
    var nullable: c.SQLSMALLINT = 0;
    const ret = c.SQLDescribeCol(
        @ptrCast(stmt),
        col,
        @ptrCast(&name_buf),
        @intCast(name_buf.len),
        &name_len,
        &data_type,
        &col_size,
        &decimal_digits,
        &nullable,
    );
    if (ret != c.SQL_SUCCESS and ret != c.SQL_SUCCESS_WITH_INFO) {
        return error.ODBCDescribeColFailed;
    }
    const n: usize = if (name_len <= 0) 0 else @min(@as(usize, @intCast(name_len)), name_buf.len);
    return allocator.dupe(u8, name_buf[0..n]);
}

/// Gets string data from the given column (1-based). Caller owns the returned slice; use allocator to free.
pub fn getDataString(allocator: std.mem.Allocator, stmt: Stmt, col: c.SQLUSMALLINT) !?[]const u8 {
    var buf: [4096]u8 = undefined;
    var len: c.SQLLEN = 0;
    const ret = c.SQLGetData(
        @ptrCast(stmt),
        col,
        c.SQL_C_CHAR,
        @ptrCast(&buf),
        @intCast(buf.len),
        &len,
    );
    if (ret == c.SQL_NO_DATA or len == c.SQL_NULL_DATA) return null;
    if (ret != c.SQL_SUCCESS and ret != c.SQL_SUCCESS_WITH_INFO) return error.ODBCGetDataFailed;
    const n: usize = if (len < 0) buf.len else @intCast(len);
    const slice = try allocator.dupe(u8, buf[0..n]);
    return slice;
}

/// High-level connection wrapper: init env, set version, connect, and run queries.
pub const Connection = struct {
    env: Env,
    dbc: Dbc,
    allocator: std.mem.Allocator,

    pub fn init(allocator: std.mem.Allocator) !Connection {
        const env = try allocEnv();
        errdefer freeEnv(env);
        try setEnvAttrOdbcVersion(env);
        const dbc = try allocDbc(env);
        errdefer freeDbc(dbc);
        return .{ .env = env, .dbc = dbc, .allocator = allocator };
    }

    pub fn deinit(self: *Connection) void {
        disconnect(self.dbc);
        freeDbc(self.dbc);
        freeEnv(self.env);
    }

    /// Connect using DSN.
    pub fn connectWithDsn(self: *Connection, dsn: []const u8, uid: ?[]const u8, pwd: ?[]const u8) !void {
        try connectDsn(self.env, self.dbc, dsn, uid, pwd);
    }

    /// Connect using a connection string.
    pub fn connectWithConnStr(self: *Connection, conn_str: []const u8) !void {
        try connectWithString(self.dbc, conn_str);
    }

    /// Execute a SQL string (e.g. "SELECT * FROM t") and optionally collect rows as strings.
    pub fn exec(self: *Connection, sql: []const u8) !void {
        const stmt = try allocStmt(self.dbc);
        defer freeStmt(stmt);
        try execDirect(stmt, sql);
    }

    /// Execute a query and iterate over rows. Caller must call row.deinit() on each row.
    pub fn query(self: *Connection, sql: []const u8) !QueryResult {
        const stmt = try allocStmt(self.dbc);
        try execDirect(stmt, sql);
        const n_cols = try numResultCols(stmt);
        return .{
            .stmt = stmt,
            .dbc = self.dbc,
            .n_cols = @intCast(n_cols),
            .allocator = self.allocator,
        };
    }
};

pub const QueryResult = struct {
    stmt: Stmt,
    dbc: Dbc,
    n_cols: u16,
    allocator: std.mem.Allocator,

    pub fn deinit(self: *QueryResult) void {
        freeStmt(self.stmt);
    }

    pub fn next(self: *QueryResult) ?Row {
        if (!fetch(self.stmt)) return null;
        return .{ .stmt = self.stmt, .n_cols = self.n_cols, .allocator = self.allocator };
    }

    pub fn columnCount(self: *QueryResult) u16 {
        return self.n_cols;
    }

    /// Get the name of the given column (1-based). Caller owns the returned slice.
    pub fn columnName(self: *QueryResult, allocator: std.mem.Allocator, col_index: u16) ![]const u8 {
        if (col_index < 1 or col_index > self.n_cols) return error.InvalidColumn;
        return getColumnName(allocator, self.stmt, @intCast(col_index));
    }
};

pub const Row = struct {
    stmt: Stmt,
    n_cols: u16,
    allocator: std.mem.Allocator,

    pub fn get(self: *const Row, col: u16) !?[]const u8 {
        if (col < 1 or col > self.n_cols) return error.InvalidColumn;
        return getDataString(self.allocator, self.stmt, @intCast(col));
    }

    pub fn deinit(self: *Row) void {
        _ = self;
    }
};
