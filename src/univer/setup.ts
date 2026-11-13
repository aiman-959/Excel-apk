/**
 * Univer bootstrap.
 *
 * Univer renders the virtualised canvas grid (only visible cells are drawn) and
 * provides editing, selection, merge, freeze, sort, filter, find/replace,
 * numfmt and conditional formatting. Its own formula execution is turned OFF
 * (`notExecuteFormula: true`): HyperFormula, running in a Web Worker, is the
 * single source of truth for computed values (see engine/formulaBridge.ts).
 *
 * Touch devices get Univer's mobile UI plugins (pan with inertia, pinch zoom,
 * tap-to-select); desktop browsers get the desktop ones.
 */
import { LocaleType, Univer, mergeLocales, UniverInstanceType } from '@univerjs/core';
import { FUniver } from '@univerjs/core/facade';
import { greenTheme } from '@univerjs/themes';
import { UniverDocsPlugin } from '@univerjs/docs';
import { UniverDocsUIPlugin } from '@univerjs/docs-ui';
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula';
import { UniverRenderEnginePlugin } from '@univerjs/engine-render';
import { UniverSheetsPlugin } from '@univerjs/sheets';
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula';
import { UniverSheetsNumfmtPlugin } from '@univerjs/sheets-numfmt';
import { UniverSheetsUIPlugin, UniverSheetsMobileUIPlugin } from '@univerjs/sheets-ui';
import { UniverUIPlugin, UniverMobileUIPlugin } from '@univerjs/ui';
import { UniverSheetsFilterPlugin } from '@univerjs/sheets-filter';
import { UniverSheetsFilterUIPlugin, UniverSheetsFilterMobileUIPlugin } from '@univerjs/sheets-filter-ui';
import { UniverSheetsSortPlugin } from '@univerjs/sheets-sort';
import { UniverSheetsSortUIPlugin, UniverSheetsSortMobileUIPlugin } from '@univerjs/sheets-sort-ui';
import { UniverFindReplacePlugin, UniverFindReplaceMobileUIPlugin } from '@univerjs/find-replace';
import { UniverSheetsFindReplacePlugin, UniverSheetsFindReplaceMobileUIPlugin } from '@univerjs/sheets-find-replace';
import { UniverSheetsConditionalFormattingPlugin } from '@univerjs/sheets-conditional-formatting';
import {
  UniverSheetsConditionalFormattingUIPlugin,
  UniverSheetsConditionalFormattingMobileUIPlugin,
} from '@univerjs/sheets-conditional-formatting-ui';

import '@univerjs/sheets/facade';
import '@univerjs/ui/facade';
import '@univerjs/docs-ui/facade';
import '@univerjs/sheets-ui/facade';
import '@univerjs/sheets-numfmt/facade';
import '@univerjs/sheets-filter/facade';
import '@univerjs/sheets-sort/facade';
import '@univerjs/sheets-find-replace/facade';
import '@univerjs/sheets-conditional-formatting/facade';

import coreEn from '@univerjs/preset-sheets-core/locales/en-US';
import filterEn from '@univerjs/preset-sheets-filter/locales/en-US';
import sortEn from '@univerjs/preset-sheets-sort/locales/en-US';
import findEn from '@univerjs/preset-sheets-find-replace/locales/en-US';
import cfEn from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';

import '@univerjs/preset-sheets-core/lib/index.css';
import '@univerjs/preset-sheets-filter/lib/index.css';
import '@univerjs/preset-sheets-sort/lib/index.css';
import '@univerjs/preset-sheets-find-replace/lib/index.css';
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css';

/** Univer's green theme, with the primary scale aligned to Petak's brand (#15803D primary, #14532D dark). */
export const petakTheme = {
  ...greenTheme,
  primary: {
    50: '#F0FDF4',
    100: '#DCFCE7',
    200: '#BBF7D0',
    300: '#86EFAC',
    400: '#4ADE80',
    500: '#22C55E',
    600: '#15803D',
    700: '#166534',
    800: '#14532D',
    900: '#052E16',
  },
};

export interface UniverHandle {
  univer: Univer;
  api: FUniver;
  mobile: boolean;
}

export function isTouchDevice(): boolean {
  return (
    typeof window !== 'undefined' &&
    (('ontouchstart' in window && navigator.maxTouchPoints > 0) || matchMedia('(pointer: coarse)').matches)
  );
}

export function createUniverApp(container: HTMLElement, opts: { mobile?: boolean } = {}): UniverHandle {
  const mobile = opts.mobile ?? isTouchDevice();
  const univer = new Univer({
    theme: petakTheme,
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: mergeLocales(coreEn, filterEn, sortEn, findEn, cfEn) },
  });

  const uiConfig = {
    container,
    header: false, // we draw our own toolbars (original design, touch-friendly)
    toolbar: false,
    footer: false,
    contextMenu: true,
  } as const;

  univer.registerPlugin(UniverDocsPlugin);
  univer.registerPlugin(UniverRenderEnginePlugin);
  univer.registerPlugin(mobile ? UniverMobileUIPlugin : UniverUIPlugin, uiConfig);
  univer.registerPlugin(UniverDocsUIPlugin);
  // Univer still parses formula strings (reference shifting on row insert, etc.)
  // but never evaluates them; HyperFormula does that in a worker.
  univer.registerPlugin(UniverFormulaEnginePlugin, { notExecuteFormula: true });
  univer.registerPlugin(UniverSheetsPlugin, { notExecuteFormula: true, onlyRegisterFormulaRelatedMutations: false as never });
  univer.registerPlugin(mobile ? UniverSheetsMobileUIPlugin : UniverSheetsUIPlugin, {
    formulaBar: false,
    footer: false,
  } as never);
  univer.registerPlugin(UniverSheetsNumfmtPlugin);
  univer.registerPlugin(UniverSheetsFormulaPlugin, { initialFormulaComputing: 2 as never /* NO_CALCULATION */ });
  univer.registerPlugin(UniverSheetsFilterPlugin);
  univer.registerPlugin(mobile ? UniverSheetsFilterMobileUIPlugin : UniverSheetsFilterUIPlugin);
  univer.registerPlugin(UniverSheetsSortPlugin);
  univer.registerPlugin(mobile ? UniverSheetsSortMobileUIPlugin : UniverSheetsSortUIPlugin);
  univer.registerPlugin(mobile ? UniverFindReplaceMobileUIPlugin : UniverFindReplacePlugin);
  univer.registerPlugin(mobile ? UniverSheetsFindReplaceMobileUIPlugin : UniverSheetsFindReplacePlugin);
  univer.registerPlugin(UniverSheetsConditionalFormattingPlugin);
  univer.registerPlugin(
    mobile ? UniverSheetsConditionalFormattingMobileUIPlugin : UniverSheetsConditionalFormattingUIPlugin,
  );

  const api = FUniver.newAPI(univer);
  return { univer, api, mobile };
}

export { UniverInstanceType };
