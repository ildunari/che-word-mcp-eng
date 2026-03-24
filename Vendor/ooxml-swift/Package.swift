// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "OOXMLSwift",
    platforms: [.macOS(.v13)],
    products: [
        .library(name: "OOXMLSwift", targets: ["OOXMLSwift"])
    ],
    dependencies: [
        .package(url: "https://github.com/weichsel/ZIPFoundation.git", from: "0.9.0")
    ],
    targets: [
        .target(
            name: "OOXMLSwift",
            dependencies: ["ZIPFoundation"]
        ),
        .testTarget(
            name: "OOXMLSwiftTests",
            dependencies: ["OOXMLSwift"]
        )
    ]
)
