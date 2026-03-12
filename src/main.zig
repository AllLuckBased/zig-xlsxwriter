const std = @import("std");
const zigxlsx = @import("zigxlsx");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    const filename = "CustCustomerV3Entity.xlsx";
    const sheet_name = "Sheet";
    std.log.info("Reading Excel file: {s}", .{filename});

    var sheet_data = try zigxlsx.read_excel(allocator, filename, sheet_name);
    defer sheet_data.deinit();

    const row = sheet_data.locRow(0); // first row
    const cell = sheet_data.locCell(2, 1); // row 2, column 1
    const by_name = sheet_data.locCellByName(2, "CustomerName");

    if (row) |r|
        std.log.info("Row 0: {d} cells", .{r.len})
    else
        std.log.info("Row 0: (null)", .{});
    if (cell) |c|
        std.log.info("Cell (2,1): {s}", .{c})
    else
        std.log.info("Cell (2,1): (null)", .{});
    if (by_name) |v|
        std.log.info("By Name: {s}", .{v})
    else
        std.log.info("By Name: (null)", .{});
}
