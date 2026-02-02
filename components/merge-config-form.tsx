'use client';

import { useState } from 'react';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card';
import { Label } from '@/components/ui/label';
import { Input } from '@/components/ui/input';
import { Checkbox } from '@/components/ui/checkbox';

export interface MergeFormConfig {
  includeTotal: boolean;
  startRow: number;
  startColumn: string;
  endColumn: string;
  includeSignature: boolean;
}

interface MergeConfigFormProps {
  config: MergeFormConfig;
  onChange: (config: MergeFormConfig) => void;
}

export function MergeConfigForm({ config, onChange }: MergeConfigFormProps) {
  const handleChange = (field: keyof MergeFormConfig, value: any) => {
    onChange({
      ...config,
      [field]: value,
    });
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle>Merge Configuration</CardTitle>
        <CardDescription>
          Configure how the Excel files should be merged
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        {/* Include Total Row */}
        <div className="space-y-2">
          <div className="flex items-center gap-2">
            <Checkbox
              id="include-total"
              checked={config.includeTotal}
              onCheckedChange={(checked) =>
                handleChange('includeTotal', checked)
              }
            />
            <Label
              htmlFor="include-total"
              className="font-medium cursor-pointer"
            >
              Include TOTAL Row
            </Label>
          </div>
          <p className="text-xs text-muted-foreground ml-6">
            Include rows containing only "TOTAL" in the output
          </p>
        </div>

        {/* Start Row */}
        <div className="space-y-2">
          <Label htmlFor="start-row" className="font-medium">
            Data Start Row
          </Label>
          <Input
            id="start-row"
            type="number"
            min="1"
            value={config.startRow}
            onChange={(e) => {
              const value = e.target.value;
              if (value === '' || value === '0') {
                return; // Don't allow empty or zero
              }
              const numValue = parseInt(value, 10);
              if (!isNaN(numValue) && numValue > 0) {
                handleChange('startRow', numValue);
              }
            }}
          />
          <p className="text-xs text-muted-foreground">
            Rows before this number will appear only once in the final file
            (treated as header)
          </p>
        </div>

        {/* Column Range */}
        <div className="grid grid-cols-2 gap-4">
          <div className="space-y-2">
            <Label htmlFor="start-column" className="font-medium">
              Start Column
            </Label>
            <Input
              id="start-column"
              type="text"
              maxLength={2}
              value={config.startColumn.toUpperCase()}
              onChange={(e) =>
                handleChange('startColumn', e.target.value.toUpperCase())
              }
              placeholder="A"
            />
            <p className="text-xs text-muted-foreground">e.g., A, B, C</p>
          </div>

          <div className="space-y-2">
            <Label htmlFor="end-column" className="font-medium">
              End Column
            </Label>
            <Input
              id="end-column"
              type="text"
              maxLength={2}
              value={config.endColumn.toUpperCase()}
              onChange={(e) =>
                handleChange('endColumn', e.target.value.toUpperCase())
              }
              placeholder="Z"
            />
            <p className="text-xs text-muted-foreground">e.g., E, F, Z</p>
          </div>
        </div>

        {/* Include Signature */}
        <div className="space-y-2">
          <div className="flex items-center gap-2">
            <Checkbox
              id="include-signature"
              checked={config.includeSignature}
              onCheckedChange={(checked) =>
                handleChange('includeSignature', checked)
              }
            />
            <Label
              htmlFor="include-signature"
              className="font-medium cursor-pointer"
            >
              Include Signature Section
            </Label>
          </div>
          <p className="text-xs text-muted-foreground ml-6">
            Include signature blocks ("Người lập phiếu" and "Người nhận") at the
            end
          </p>
        </div>
      </CardContent>
    </Card>
  );
}
