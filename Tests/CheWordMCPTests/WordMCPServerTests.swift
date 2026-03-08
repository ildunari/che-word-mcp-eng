import XCTest
@testable import CheWordMCP

final class WordMCPServerTests: XCTestCase {
    func testServerInitializes() async {
        let server = await WordMCPServer()
        XCTAssertNotNil(server)
    }
}
