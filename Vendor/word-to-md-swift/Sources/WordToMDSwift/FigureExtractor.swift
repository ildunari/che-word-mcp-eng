import Foundation
import OOXMLSwift

#if canImport(AppKit)
import AppKit
#endif

/// 圖片提取器（Tier 2+）
///
/// 負責將 WordDocument 中的 ImageReference（含 binary data）寫入 figures 目錄。
/// 回傳相對路徑供 Markdown 引用：`![alt](figures/image1.png)`
///
/// Practical Mode（preserveOriginalFormat = false）：
/// EMF/WMF/TIFF/BMP 等非 web-friendly 格式自動轉為 PNG。
struct FigureExtractor {
    let directory: URL
    private var extractedIds: Set<String> = []

    init(directory: URL) {
        self.directory = directory
    }

    /// 建立 figures 目錄（如果不存在）
    func createDirectory() throws {
        try FileManager.default.createDirectory(
            at: directory,
            withIntermediateDirectories: true
        )
    }

    /// 提取圖片到 figures 目錄
    /// - Parameters:
    ///   - imageRef: 圖片參照
    ///   - preserveOriginalFormat: true 時保留原始格式，false 時非 web-friendly 格式轉 PNG
    /// - Returns: 相對路徑（如 `figures/image1.png`）
    mutating func extract(_ imageRef: ImageReference, preserveOriginalFormat: Bool = false) throws -> String {
        let outputName = outputFileName(imageRef, preserveOriginal: preserveOriginalFormat)

        // 避免重複提取
        guard !extractedIds.contains(imageRef.id) else {
            return relativePath(for: outputName)
        }

        let data: Data
        let fileName: String

        if preserveOriginalFormat || isWebFriendly(imageRef.fileName) {
            data = imageRef.data
            fileName = imageRef.fileName
        } else {
            // EMF/WMF/TIFF/BMP → PNG
            data = try convertToPNG(imageRef.data)
            fileName = replaceExtension(imageRef.fileName, with: "png")
        }

        let fileURL = directory.appendingPathComponent(fileName)
        try data.write(to: fileURL)
        extractedIds.insert(imageRef.id)

        return relativePath(for: fileName)
    }

    // MARK: - Private

    private func relativePath(for fileName: String) -> String {
        "figures/\(fileName)"
    }

    /// 判斷副檔名是否為 web-friendly 格式
    private func isWebFriendly(_ fileName: String) -> Bool {
        let ext = (fileName as NSString).pathExtension.lowercased()
        let webFriendly: Set<String> = ["png", "jpg", "jpeg", "gif", "svg", "webp"]
        return webFriendly.contains(ext)
    }

    /// 決定輸出檔名
    private func outputFileName(_ imageRef: ImageReference, preserveOriginal: Bool) -> String {
        if preserveOriginal || isWebFriendly(imageRef.fileName) {
            return imageRef.fileName
        }
        return replaceExtension(imageRef.fileName, with: "png")
    }

    /// 替換副檔名
    private func replaceExtension(_ fileName: String, with newExt: String) -> String {
        let name = (fileName as NSString).deletingPathExtension
        return "\(name).\(newExt)"
    }

    /// 將圖片資料轉為 PNG（使用 AppKit）
    private func convertToPNG(_ data: Data) throws -> Data {
        #if canImport(AppKit)
        guard let image = NSImage(data: data) else {
            throw FigureExtractorError.conversionFailed("Cannot create NSImage from data")
        }
        guard let tiffData = image.tiffRepresentation,
              let bitmap = NSBitmapImageRep(data: tiffData),
              let pngData = bitmap.representation(using: .png, properties: [:]) else {
            throw FigureExtractorError.conversionFailed("Cannot convert image to PNG")
        }
        return pngData
        #else
        throw FigureExtractorError.conversionFailed("AppKit not available — cannot convert image to PNG")
        #endif
    }
}

/// 圖片提取錯誤
enum FigureExtractorError: Error, LocalizedError {
    case conversionFailed(String)

    var errorDescription: String? {
        switch self {
        case .conversionFailed(let message):
            return "Image conversion failed: \(message)"
        }
    }
}
