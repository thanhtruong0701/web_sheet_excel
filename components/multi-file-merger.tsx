'use client';

import { useState, useEffect } from 'react';
import { Download, Loader2, FileSpreadsheet, ChevronDown, ChevronUp } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { FileUploader } from '@/components/file-uploader';
import { MergeConfigForm, type MergeFormConfig } from '@/components/merge-config-form';
import { useToast } from '@/hooks/use-toast';
import { mergeMultipleFiles, readSheetNames } from '@/lib/excel-utils';

interface FileWithSheets {
  file: File;
  sheets: string[];
  loading: boolean;
}

const DEFAULT_CONFIG: MergeFormConfig = {
  includeTotal: true,
  startRow: 2,
  startColumn: 'A',
  endColumn: 'Z',
  includeSignature: true,
};

export function MultiFileMerger() {
  const [filesWithSheets, setFilesWithSheets] = useState<FileWithSheets[]>([]);
  const [config, setConfig] = useState<MergeFormConfig>(DEFAULT_CONFIG);
  const [loading, setLoading] = useState(false);
  const [expandedFiles, setExpandedFiles] = useState<Set<number>>(new Set());
  const { toast } = useToast();

  const files = filesWithSheets.map(f => f.file);

  const handleFilesChange = async (newFiles: File[]) => {
    // Find newly added files (files that aren't already in our list)
    const existingNames = new Set(filesWithSheets.map(f => f.file.name + f.file.size));
    const addedFiles = newFiles.filter(f => !existingNames.has(f.name + f.size));

    // Create entries for new files
    const newEntries: FileWithSheets[] = addedFiles.map(file => ({
      file,
      sheets: [],
      loading: true,
    }));

    // Keep existing entries that are still in newFiles
    const newFileKeys = new Set(newFiles.map(f => f.name + f.size));
    const keptEntries = filesWithSheets.filter(f => newFileKeys.has(f.file.name + f.file.size));

    const updatedList = [...keptEntries, ...newEntries];
    setFilesWithSheets(updatedList);

    // Read sheet names for new files
    for (let i = 0; i < newEntries.length; i++) {
      const entry = newEntries[i];
      try {
        const sheets = await readSheetNames(entry.file);
        setFilesWithSheets(prev =>
          prev.map(f =>
            f.file.name === entry.file.name && f.file.size === entry.file.size
              ? { ...f, sheets, loading: false }
              : f
          )
        );
      } catch {
        setFilesWithSheets(prev =>
          prev.map(f =>
            f.file.name === entry.file.name && f.file.size === entry.file.size
              ? { ...f, sheets: ['(Không đọc được)'], loading: false }
              : f
          )
        );
      }
    }
  };

  const handleFilesChangeWrapper = (newFiles: File[]) => {
    handleFilesChange(newFiles);
  };

  const toggleExpanded = (index: number) => {
    setExpandedFiles(prev => {
      const next = new Set(prev);
      if (next.has(index)) {
        next.delete(index);
      } else {
        next.add(index);
      }
      return next;
    });
  };

  const totalSheets = filesWithSheets.reduce((sum, f) => sum + f.sheets.length, 0);

  const handleMerge = async () => {
    if (files.length === 0) {
      toast({
        title: 'Chưa chọn file',
        description: 'Vui lòng chọn ít nhất 1 file Excel để gộp.',
        variant: 'destructive',
      });
      return;
    }

    setLoading(true);
    try {
      const buffer = await mergeMultipleFiles(files, config);

      const timestamp = new Date().toISOString().split('T')[0];
      const fileName = `gop_tat_ca_${timestamp}.xlsx`;

      const blob = new Blob([new Uint8Array(buffer)], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      toast({
        title: 'Thành công!',
        description: `Đã gộp ${files.length} file (${totalSheets} sheet) thành 1 file. Đang tải xuống...`,
      });

      setFilesWithSheets([]);
      setExpandedFiles(new Set());
    } catch (error) {
      console.error('Error:', error);
      toast({
        title: 'Lỗi',
        description:
          error instanceof Error
            ? error.message
            : 'Không thể gộp file. Vui lòng thử lại.',
        variant: 'destructive',
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* File Uploader */}
      <FileUploader files={files} onFilesChange={handleFilesChangeWrapper} />

      {/* Sheet Preview */}
      {filesWithSheets.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <FileSpreadsheet className="h-5 w-5" />
              Xem trước Sheet
            </CardTitle>
            <CardDescription>
              Tổng cộng {filesWithSheets.length} file — {totalSheets} sheet sẽ được gộp + 1 sheet Tổng hợp
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-2">
            {filesWithSheets.map((entry, fileIndex) => (
              <div
                key={`${entry.file.name}-${entry.file.size}`}
                className="border border-border rounded-lg overflow-hidden"
              >
                <button
                  onClick={() => toggleExpanded(fileIndex)}
                  className="w-full flex items-center justify-between p-3 hover:bg-muted/50 transition-colors text-left"
                >
                  <div className="flex items-center gap-2">
                    <FileSpreadsheet className="h-4 w-4 text-green-600" />
                    <span className="text-sm font-medium">{entry.file.name}</span>
                    {entry.loading ? (
                      <Loader2 className="h-3 w-3 animate-spin text-muted-foreground" />
                    ) : (
                      <span className="text-xs text-muted-foreground bg-muted px-2 py-0.5 rounded-full">
                        {entry.sheets.length} sheet
                      </span>
                    )}
                  </div>
                  {expandedFiles.has(fileIndex) ? (
                    <ChevronUp className="h-4 w-4 text-muted-foreground" />
                  ) : (
                    <ChevronDown className="h-4 w-4 text-muted-foreground" />
                  )}
                </button>
                {expandedFiles.has(fileIndex) && !entry.loading && (
                  <div className="px-3 pb-3 pt-0">
                    <div className="space-y-1 ml-6">
                      {entry.sheets.map((sheetName, sheetIndex) => (
                        <div
                          key={sheetIndex}
                          className="flex items-center gap-2 text-sm text-muted-foreground"
                        >
                          <div className="w-1.5 h-1.5 rounded-full bg-primary/50" />
                          {sheetName}
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            ))}
          </CardContent>
        </Card>
      )}

      {/* Config Form */}
      <MergeConfigForm config={config} onChange={setConfig} />

      {/* Action Buttons */}
      <div className="flex gap-2 justify-end">
        <Button
          variant="outline"
          onClick={() => {
            setFilesWithSheets([]);
            setExpandedFiles(new Set());
            setConfig(DEFAULT_CONFIG);
          }}
          disabled={loading}
        >
          Reset
        </Button>
        <Button onClick={handleMerge} disabled={loading || files.length === 0}>
          {loading ? (
            <>
              <Loader2 className="mr-2 h-4 w-4 animate-spin" />
              Đang gộp...
            </>
          ) : (
            <>
              <Download className="mr-2 h-4 w-4" />
              Gộp tất cả & Tải xuống
            </>
          )}
        </Button>
      </div>

      {/* Info Section */}
      <div className="bg-muted/50 border border-border rounded-lg p-4 space-y-2">
        <h3 className="font-semibold text-sm">Cách hoạt động:</h3>
        <ul className="text-sm text-muted-foreground space-y-1 ml-4">
          <li>• Import nhiều file Excel cùng lúc</li>
          <li>• Tất cả sheet gốc từ mỗi file được giữ nguyên trong file output</li>
          <li>• Sheet trùng tên sẽ được đánh số tự động (Sheet1, Sheet1 (2)...)</li>
          <li>
            • Tạo thêm 1 sheet <strong>&quot;Tổng hợp&quot;</strong> gộp dữ liệu từ tất cả sheet
          </li>
          <li>• Sheet Tổng hợp sử dụng cùng cấu hình merge bên dưới</li>
        </ul>
      </div>
    </div>
  );
}
