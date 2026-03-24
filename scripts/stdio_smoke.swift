#!/usr/bin/env swift

import Foundation

struct JSONRPCClient {
    let process: Process
    let stdin: FileHandle
    let stdout: FileHandle

    init(binaryPath: String) throws {
        process = Process()
        process.executableURL = URL(fileURLWithPath: binaryPath)
        let stdinPipe = Pipe()
        let stdoutPipe = Pipe()
        let stderrPipe = Pipe()
        process.standardInput = stdinPipe
        process.standardOutput = stdoutPipe
        process.standardError = stderrPipe
        try process.run()
        stdin = stdinPipe.fileHandleForWriting
        stdout = stdoutPipe.fileHandleForReading
    }

    func send(_ message: [String: Any]) throws {
        let data = try JSONSerialization.data(withJSONObject: message)
        stdin.write(data)
        stdin.write(Data([0x0a]))
    }

    func receive() throws -> [String: Any] {
        var data = Data()
        while true {
            let chunk = try stdout.read(upToCount: 1) ?? Data()
            if chunk.isEmpty {
                throw NSError(domain: "stdio_smoke", code: 1, userInfo: [NSLocalizedDescriptionKey: "EOF while waiting for JSON-RPC response"])
            }
            if chunk[0] == 0x0a { break }
            data.append(chunk)
        }
        let object = try JSONSerialization.jsonObject(with: data, options: [])
        guard let dict = object as? [String: Any] else {
            throw NSError(domain: "stdio_smoke", code: 2, userInfo: [NSLocalizedDescriptionKey: "Invalid JSON-RPC payload"])
        }
        return dict
    }

    func call(id: Int, name: String, arguments: [String: Any] = [:]) throws -> String {
        try send([
            "jsonrpc": "2.0",
            "id": id,
            "method": "tools/call",
            "params": [
                "name": name,
                "arguments": arguments
            ]
        ])
        let response = try receive()
        guard
            let result = response["result"] as? [String: Any],
            let content = result["content"] as? [[String: Any]],
            let text = content.first?["text"] as? String
        else {
            throw NSError(domain: "stdio_smoke", code: 3, userInfo: [NSLocalizedDescriptionKey: "Malformed tools/call response for \(name)"])
        }
        return text
    }
}

@discardableResult
func run(_ command: [String]) throws -> String {
    let process = Process()
    process.executableURL = URL(fileURLWithPath: command[0])
    process.arguments = Array(command.dropFirst())
    let pipe = Pipe()
    process.standardOutput = pipe
    process.standardError = pipe
    try process.run()
    process.waitUntilExit()
    let data = pipe.fileHandleForReading.readDataToEndOfFile()
    let output = String(decoding: data, as: UTF8.self)
    guard process.terminationStatus == 0 else {
        throw NSError(domain: "stdio_smoke", code: Int(process.terminationStatus), userInfo: [NSLocalizedDescriptionKey: output])
    }
    return output
}

func require(_ condition: @autoclosure () -> Bool, _ message: String) throws {
    if !condition() {
        throw NSError(domain: "stdio_smoke", code: 4, userInfo: [NSLocalizedDescriptionKey: message])
    }
}

let repoRoot = URL(fileURLWithPath: FileManager.default.currentDirectoryPath)
let env = ProcessInfo.processInfo.environment
let binaryPath: String
if CommandLine.arguments.count > 1 {
    binaryPath = CommandLine.arguments[1]
} else if let envPath = env["CHE_WORD_MCP_BINARY"] {
    binaryPath = envPath
} else {
    let binPath = try run([
        "/usr/bin/env",
        "swift",
        "build",
        "--show-bin-path"
    ]).trimmingCharacters(in: .whitespacesAndNewlines)
    binaryPath = URL(fileURLWithPath: binPath).appendingPathComponent("CheWordMCP").path
}
let docPath = "/tmp/che-word-stdio-smoke.docx"

try? FileManager.default.removeItem(atPath: docPath)

let client = try JSONRPCClient(binaryPath: binaryPath)
defer {
    client.process.terminate()
}

try client.send([
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": [
        "protocolVersion": "2025-03-26",
        "capabilities": [:],
        "clientInfo": [
            "name": "stdio-smoke",
            "version": "1.0"
        ]
    ]
])
_ = try client.receive()
try client.send([
    "jsonrpc": "2.0",
    "method": "notifications/initialized"
])

let create = try client.call(id: 2, name: "create_document", arguments: ["doc_id": "doc"])
let insertParagraph = try client.call(id: 3, name: "insert_paragraph", arguments: ["doc_id": "doc", "text": "Smoke paragraph"])
let insertComment = try client.call(id: 4, name: "insert_comment", arguments: ["doc_id": "doc", "paragraph_index": 0, "author": "Smoke", "text": "Please revise"])
let finalize = try client.call(id: 5, name: "finalize_document", arguments: ["doc_id": "doc", "path": docPath])

try require(create.contains("Created new document"), "create_document did not succeed")
try require(insertParagraph.contains("Inserted paragraph"), "insert_paragraph did not succeed")
try require(insertComment.contains("Inserted comment"), "insert_comment did not succeed")
try require(finalize.contains("Finalized document"), "finalize_document did not succeed")

let baselineOpen = try client.call(id: 6, name: "open_document", arguments: ["path": docPath, "doc_id": "baseline", "autosave": true])
let baselineAccept = try client.call(id: 7, name: "accept_all_revisions", arguments: ["doc_id": "baseline"])
let baselineClose = try client.call(id: 8, name: "close_document", arguments: ["doc_id": "baseline"])

try require(baselineOpen.contains("Opened document"), "baseline open_document did not succeed")
try require(baselineAccept.contains("Accepted"), "baseline accept_all_revisions did not succeed")
try require(baselineClose.contains("Closed document"), "baseline close_document did not succeed")

let open = try client.call(id: 9, name: "open_document", arguments: ["path": docPath, "doc_id": "edit", "autosave": true])
let reply = try client.call(id: 10, name: "reply_to_comment", arguments: ["doc_id": "edit", "parent_comment_id": 1, "text": "Reply from harness", "author": "Smoke"])
let update = try client.call(id: 11, name: "update_paragraph", arguments: ["doc_id": "edit", "index": 0, "text": "Edited by stdio smoke harness"])
let revisions = try client.call(id: 12, name: "get_revisions", arguments: ["doc_id": "edit"])
let rejectAll = try client.call(id: 13, name: "reject_all_revisions", arguments: ["doc_id": "edit"])
let closeAfterReject = try client.call(id: 14, name: "close_document", arguments: ["doc_id": "edit"])

try require(open.contains("Opened document"), "open_document did not succeed")
try require(reply.contains("reply ID"), "reply_to_comment did not succeed")
try require(update.contains("Updated paragraph"), "update_paragraph did not succeed")
try require(revisions.contains("Revisions in document"), "get_revisions did not report native revisions")
try require(rejectAll.contains("Rejected"), "reject_all_revisions did not succeed")
try require(closeAfterReject.contains("Closed document"), "close_document after reject did not succeed")

let reopen = try client.call(id: 15, name: "open_document", arguments: ["path": docPath, "doc_id": "accept", "autosave": true])
let secondUpdate = try client.call(id: 16, name: "update_paragraph", arguments: ["doc_id": "accept", "index": 0, "text": "Accepted by stdio smoke harness"])
let acceptAll = try client.call(id: 17, name: "accept_all_revisions", arguments: ["doc_id": "accept"])
let save = try client.call(id: 18, name: "save_document", arguments: ["doc_id": "accept"])
let comments = try client.call(id: 19, name: "list_comments", arguments: ["doc_id": "accept"])
let close = try client.call(id: 20, name: "close_document", arguments: ["doc_id": "accept"])

try require(reopen.contains("Opened document"), "second open_document did not succeed")
try require(secondUpdate.contains("Updated paragraph"), "second update_paragraph did not succeed")
try require(acceptAll.contains("Accepted"), "accept_all_revisions did not succeed")
try require(save.contains("Saved document"), "save_document did not succeed")
try require(comments.contains("reply to 1"), "list_comments did not expose reply threading")
try require(close.contains("Closed document"), "final close_document did not succeed")

let documentXML = try run(["/usr/bin/unzip", "-p", docPath, "word/document.xml"])
let settingsXML = try run(["/usr/bin/unzip", "-p", docPath, "word/settings.xml"])
let commentsXML = try run(["/usr/bin/unzip", "-p", docPath, "word/comments.xml"])
let commentsExtendedXML = try run(["/usr/bin/unzip", "-p", docPath, "word/commentsExtended.xml"])

try require(documentXML.contains("Accepted by stdio smoke harness"), "document.xml does not contain accepted text")
try require(settingsXML.contains("trackRevisions"), "settings.xml is missing w:trackRevisions")
try require(commentsXML.contains("Reply from harness"), "comments.xml is missing the reply text")
try require(commentsExtendedXML.contains("paraIdParent"), "commentsExtended.xml is missing reply threading metadata")

print("stdio smoke passed")
