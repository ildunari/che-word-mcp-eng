// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "CheWordMCP",
    platforms: [.macOS(.v13)],
    dependencies: [
        .package(path: "Vendor/swift-sdk"),
        .package(path: "Vendor/ooxml-swift"),
        .package(path: "Vendor/word-to-md-swift"),
    ],
    targets: [
        .executableTarget(
            name: "CheWordMCP",
            dependencies: [
                .product(name: "MCP", package: "swift-sdk"),
                .product(name: "OOXMLSwift", package: "ooxml-swift"),
                .product(name: "WordToMDSwift", package: "word-to-md-swift"),
            ]
        ),
        .testTarget(
            name: "CheWordMCPTests",
            dependencies: ["CheWordMCP"]
        )
    ]
)
