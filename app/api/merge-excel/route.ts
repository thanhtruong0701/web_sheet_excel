import { NextRequest, NextResponse } from 'next/server';
import { mergeExcelFiles } from '@/lib/excel-utils';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();

    const files: File[] = [];
    let config = {
      includeTotal: false,
      startRow: 1,
      startColumn: 'A',
      endColumn: 'Z',
      includeSignature: false,
    };

    // Extract files
    const fileEntries = formData.getAll('files');
    for (const entry of fileEntries) {
      if (entry instanceof File) {
        files.push(entry);
      }
    }

    // Extract config
    const configData = formData.get('config');
    if (configData) {
      config = JSON.parse(configData as string);
    }

    if (files.length === 0) {
      return NextResponse.json(
        { error: 'No files provided' },
        { status: 400 }
      );
    }

    const buffer = await mergeExcelFiles(files, config);

    const firstFile = files[0];
    const timestamp = new Date().toISOString().split('T')[0];
    const originalName = firstFile?.name || `${timestamp}.xlsx`;
    const fileName = originalName.toLowerCase().endsWith('.xlsx')
      ? `merged_${originalName}`
      : `merged_${originalName}.xlsx`;

    return new NextResponse(new Uint8Array(buffer), {
      headers: {
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${fileName}"`,
      },
    });
  } catch (error) {
    console.error('Error merging files:', error);
    return NextResponse.json(
      { error: 'Failed to merge files' },
      { status: 500 }
    );
  }
}
