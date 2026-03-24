// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "WordToMDSwift",
    platforms: [.macOS(.v13)],
    products: [
        .library(name: "WordToMDSwift", targets: ["WordToMDSwift"])
    ],
    dependencies: [
        .package(path: "../doc-converter-swift"),
        .package(path: "../ooxml-swift"),
        .package(url: "https://github.com/kiki830621/markdown-swift.git", from: "0.1.0"),
    ],
    targets: [
        .target(
            name: "WordToMDSwift",
            dependencies: [
                .product(name: "DocConverterSwift", package: "doc-converter-swift"),
                .product(name: "OOXMLSwift", package: "ooxml-swift"),
                .product(name: "MarkdownSwift", package: "markdown-swift"),
            ]
        ),
        .testTarget(
            name: "WordToMDSwiftTests",
            dependencies: ["WordToMDSwift"]
        )
    ]
)
