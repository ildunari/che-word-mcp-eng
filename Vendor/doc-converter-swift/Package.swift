// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "DocConverterSwift",
    platforms: [.macOS(.v13)],
    products: [
        .library(name: "DocConverterSwift", targets: ["DocConverterSwift"])
    ],
    targets: [
        .target(name: "DocConverterSwift", path: "Sources/CommonConverterSwift"),
        .testTarget(
            name: "DocConverterSwiftTests",
            dependencies: ["DocConverterSwift"]
        )
    ]
)
