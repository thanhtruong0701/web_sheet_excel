'use client';

import { useState } from 'react';
import { Download, Loader2 } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { FileUploader } from '@/components/file-uploader';
import { MergeConfigForm, type MergeFormConfig } from '@/components/merge-config-form';
import { useToast } from '@/hooks/use-toast';

const DEFAULT_CONFIG: MergeFormConfig = {
  includeTotal: true,
  startRow: 2,
  startColumn: 'A',
  endColumn: 'Z',
  includeSignature: true,
};

export default function Home() {
  const [files, setFiles] = useState<File[]>([]);
  const [config, setConfig] = useState<MergeFormConfig>(DEFAULT_CONFIG);
  const [loading, setLoading] = useState(false);
  const { toast } = useToast();

  const handleMerge = async () => {
    if (files.length === 0) {
      toast({
        title: 'No files selected',
        description: 'Please select at least one Excel file to merge.',
        variant: 'destructive',
      });
      return;
    }

    setLoading(true);
    try {
      const formData = new FormData();

      // Add files
      files.forEach((file) => {
        formData.append('files', file);
      });

      // Add config
      formData.append('config', JSON.stringify(config));

      const response = await fetch('/api/merge-excel', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        throw new Error('Failed to merge files');
      }

      // Download the file
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'consolidated.xlsx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      toast({
        title: 'Success',
        description: 'Files merged successfully and download started.',
      });

      // Reset only files, keep config for next merge
      setFiles([]);
    } catch (error) {
      console.error('Error:', error);
      toast({
        title: 'Error',
        description:
          error instanceof Error
            ? error.message
            : 'Failed to merge files. Please try again.',
        variant: 'destructive',
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-background py-8 px-4">
      <div className="max-w-4xl mx-auto space-y-6">
        {/* Header */}
        <div className="space-y-2">
          <h1 className="text-3xl font-bold tracking-tight">Excel Sheet Merger</h1>
          <p className="text-muted-foreground">
            Merge multiple Excel files into one consolidated spreadsheet with
            customizable options
          </p>
        </div>

        {/* File Uploader */}
        <FileUploader files={files} onFilesChange={setFiles} />

        {/* Config Form */}
        <MergeConfigForm config={config} onChange={setConfig} />

        {/* Action Buttons */}
        <div className="flex gap-2 justify-end">
          <Button
            variant="outline"
            onClick={() => {
              setFiles([]);
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
                Merging...
              </>
            ) : (
              <>
                <Download className="mr-2 h-4 w-4" />
                Merge & Download
              </>
            )}
          </Button>
        </div>

        {/* Info Section */}
        <div className="bg-muted/50 border border-border rounded-lg p-4 space-y-2">
          <h3 className="font-semibold text-sm">How it works:</h3>
          <ul className="text-sm text-muted-foreground space-y-1 ml-4">
            <li>• Upload one or more Excel files</li>
            <li>
              • Rows before the "Data Start Row" are treated as headers and appear
              only once
            </li>
            <li>• Data rows from all files are combined in sequence</li>
            <li>
              • TOTAL rows can be included to show subtotals for each file
            </li>
            <li>
              • Signature sections are added at the end if enabled
            </li>
            <li>• Only specified columns are included in the output</li>
          </ul>
        </div>
      </div>
    </div>
  );
}
