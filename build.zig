const std = @import("std");

const dlls = [_][]const u8{
    "vendor/xlsxio/bin/xlsxio_read.dll",
    "vendor/xlsxio/bin/xlsxio_write.dll",
    "vendor/xlsxio/bin/libexpat.dll",
    "vendor/xlsxio/bin/minizip.dll",
    "vendor/xlsxio/bin/zlib1.dll",
    "vendor/xlsxio/bin/bz2.dll",
};

/// Helper function for consumers to install necessary DLLs.
pub fn installDlls(b: *std.Build, artifact: *std.Build.Step.Compile) void {
    for (dlls) |dll_path| {
        const dll_basename = std.fs.path.basename(dll_path);
        b.installFile(dll_path, dll_basename);
    }
    if (artifact.kind == .exe) {
        const run_cmd = b.addRunArtifact(artifact);
        run_cmd.addPathDir("bin");
    }
}

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    // Expose an option "install_dlls" (default true)
    const install_dlls_opt = b.option(bool, "install_dlls", "Install DLLs") orelse true;

    // Define the xlsxio module.
    const xlsxio_mod = b.addModule("xlsxio", .{
        .root_source_file = b.path("src/xlsxio.zig"),
        .target = target,
        .optimize = optimize,
    });
    // Zig 0.16: libc linking is controlled per-module.
    // Required so @cImport can use libc headers like stdlib.h.
    xlsxio_mod.link_libc = true;
    // Add the include path so that @cImport finds xlsxio_read.h
    xlsxio_mod.addIncludePath(b.path("vendor/xlsxio/include"));

    // Add the build helper module.
    _ = b.addModule("xlsxio_build", .{
        .root_source_file = b.path("src/build_module.zig"),
        .target = target,
        .optimize = optimize,
    });

    // ODBC module (unixODBC on Unix, Windows ODBC on Windows)
    const odbc_mod = b.addModule("odbc", .{
        .root_source_file = b.path("src/odbc.zig"),
        .target = target,
        .optimize = optimize,
    });
    odbc_mod.link_libc = true;
    odbc_mod.addIncludePath(b.path("vendor/unixODBC/include"));
    const is_windows = target.result.os.tag == .windows;
    if (is_windows) {
        odbc_mod.linkSystemLibrary("odbc32", .{});
    } else {
        odbc_mod.linkSystemLibrary("odbc", .{});
    }

    // Create a shared library that consumers can link against.
    const lib = b.addLibrary(.{
        .name = "zig_xlsxio",
        .linkage = .static,
        .root_module = b.createModule(.{
            .target = target,
            .optimize = optimize,
        }),
    });

    // Zig 0.16: libc linking is controlled per-module.
    lib.root_module.link_libc = true;

    // Create a private `xlsxio` module (already created as `xlsxio_mod`) and
    // a public `zigxlsx` module that imports `xlsxio` privately. This makes
    // `zigxlsx` the public API while keeping `xlsxio` internal to this package.
    const zigxlsx_mod = b.createModule(.{
        .root_source_file = b.path("src/zigxlsx.zig"),
        .target = target,
        .optimize = optimize,
        .link_libc = true,
        .imports = &.{
            .{ .name = "xlsxio", .module = xlsxio_mod },
            .{ .name = "odbc", .module = odbc_mod },
        },
    });

    // Export the `zigxlsx` and `odbc` modules to consumers.
    lib.root_module.addImport("zigxlsx", zigxlsx_mod);
    lib.root_module.addImport("odbc", odbc_mod);

    const xlsxio_include = b.path("vendor/xlsxio/include");
    const xlsxio_lib = b.path("vendor/xlsxio/lib");
    lib.root_module.addIncludePath(xlsxio_include);
    lib.root_module.addLibraryPath(xlsxio_lib);
    lib.root_module.addObjectFile(b.path("vendor/xlsxio/lib/xlsxio_read.lib"));
    lib.root_module.addObjectFile(b.path("vendor/xlsxio/lib/xlsxio_write.lib"));

    // Conditionally install DLLs if the option is true.
    if (install_dlls_opt) {
        for (dlls) |dll_path| {
            const dll_basename = std.fs.path.basename(dll_path);
            b.installFile(dll_path, dll_basename);
        }
    }

    b.installArtifact(lib);

    const options = b.addOptions();
    options.addOption([]const u8, "include_path", "vendor/xlsxio/include");
    options.addOption([]const u8, "lib_path", "vendor/xlsxio/lib");
    options.addOption([]const []const u8, "system_libs", &[_][]const u8{ "xlsxio_read", "xlsxio_write" });
    options.addOption(bool, "link_libc", true);

    // Create a test artifact.
    const test_mod = b.createModule(.{
        .root_source_file = b.path("src/xlsxio.zig"),
        .target = target,
        .optimize = optimize,
    });
    test_mod.link_libc = true;

    // Create the test artifact using the module
    const tests = b.addTest(.{
        .root_module = test_mod,
    });

    test_mod.addIncludePath(b.path("vendor/xlsxio/include"));
    test_mod.addLibraryPath(b.path("vendor/xlsxio/lib"));
    test_mod.addObjectFile(b.path("vendor/xlsxio/lib/xlsxio_read.lib"));
    test_mod.addObjectFile(b.path("vendor/xlsxio/lib/xlsxio_write.lib"));

    const run_tests = b.addRunArtifact(tests);
    run_tests.addPathDir("vendor/xlsxio/bin");

    const test_step = b.step("test", "Run library tests");
    test_step.dependOn(&run_tests.step);

    // Create a demo executable to show usage.
    const exe_mod = b.createModule(.{
        .root_source_file = b.path("src/main.zig"),
        .target = target,
        .optimize = optimize,
    });
    exe_mod.link_libc = true;

    exe_mod.addIncludePath(xlsxio_include);
    exe_mod.addLibraryPath(xlsxio_lib);
    exe_mod.addObjectFile(b.path("vendor/xlsxio/lib/xlsxio_read.lib"));
    exe_mod.addObjectFile(b.path("vendor/xlsxio/lib/xlsxio_write.lib"));

    // Create the executable using the module
    const exe = b.addExecutable(.{
        .name = "xlsxio_demo",
        .root_module = exe_mod,
    });

    // Attach the Zig modules
    exe.root_module.addImport("xlsxio", xlsxio_mod);
    exe.root_module.addImport("zigxlsx", zigxlsx_mod);

    // Install the executable
    b.installArtifact(exe);

    const run_cmd = b.addRunArtifact(exe);
    run_cmd.addPathDir("vendor/xlsxio/bin");
    if (b.args) |args| {
        run_cmd.addArgs(args);
    }
    const run_step = b.step("run", "Run the demo application");
    run_step.dependOn(&run_cmd.step);

    // ODBC integration test executable
    const odbc_exe_mod = b.createModule(.{
        .root_source_file = b.path("src/odbc_demo.zig"),
        .target = target,
        .optimize = optimize,
    });
    odbc_exe_mod.link_libc = true;
    odbc_exe_mod.addIncludePath(xlsxio_include);
    odbc_exe_mod.addLibraryPath(xlsxio_lib);
    odbc_exe_mod.addObjectFile(b.path("vendor/xlsxio/lib/xlsxio_read.lib"));
    odbc_exe_mod.addObjectFile(b.path("vendor/xlsxio/lib/xlsxio_write.lib"));

    const odbc_exe = b.addExecutable(.{
        .name = "odbc_demo",
        .root_module = odbc_exe_mod,
    });
    odbc_exe.root_module.addImport("odbc", odbc_mod);
    odbc_exe.root_module.addImport("zigxlsx", zigxlsx_mod);
    b.installArtifact(odbc_exe);

    const run_odbc_cmd = b.addRunArtifact(odbc_exe);
    run_odbc_cmd.addPathDir("vendor/xlsxio/bin");
    if (b.args) |args| {
        run_odbc_cmd.addArgs(args);
    }
    const run_odbc_step = b.step("run-odbc", "Run the ODBC demo (usage: zig build run-odbc -- <conn_str> [query])");
    run_odbc_step.dependOn(&run_odbc_cmd.step);

    // Optionally add a step to install DLLs only.
    if (install_dlls_opt) {
        const install_dlls_step = b.step("install-dlls", "Install DLLs only");
        for (dlls) |dll_path| {
            const dll_basename = std.fs.path.basename(dll_path);
            const install_file_step = b.addInstallFile(b.path(dll_path), dll_basename);
            install_dlls_step.dependOn(&install_file_step.step);
        }
    }
}
