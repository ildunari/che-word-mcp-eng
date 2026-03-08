// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "CheWordMCP",
    platforms: [.macOS(.v13)],
    dependencies: [
        .package(path: "Vendor/swift-sdk"),
        .package(url: "https://github.com/kiki830621/ooxml-swift.git", from: "0.5.1"),
        .package(url: "https://github.com/kiki830621/word-to-md-swift.git", from: "0.3.0"),
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
