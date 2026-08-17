import type {
  AddChartFromSelectionResult,
  BasicPresentationFacade,
  DocsRangeFacade,
  DocsRangeValue,
  DocsTableFacade,
  EditorTextFormat,
  OfficeSDK,
  PresentationFacade,
  PresentationSlideFacade,
  PresentationShape,
  PresentationShapeBase,
  PresentationTableCell,
  PresentationTableItem,
  PresentationTableRange,
  PresentationTableSelection,
  PresentationTextShape,
  PresentationTextRangeFacade,
  PresentationTextRangeValue,
  SheetRangeValue,
  SheetRangeFacade,
  SheetSelection
} from './OfficeSDK'

type IsAssignable<T, U> = T extends U ? true : false
type IsEqual<T, U> = (<V>() => V extends T ? 1 : 2) extends <V>() => V extends U
  ? 1
  : 2
  ? (<V>() => V extends U ? 1 : 2) extends <V>() => V extends T ? 1 : 2
    ? true
    : false
  : false
type Assert<T extends true> = T
type RootSelection = NonNullable<OfficeSDK['selection']>
type DocsSelection = Extract<
  RootSelection,
  { getRange: (value?: DocsRangeValue) => Promise<DocsRangeFacade | null> }
>
type PresentationSelection = Extract<
  RootSelection,
  {
    getTextRange: (
      value?: PresentationTextRangeValue
    ) => Promise<PresentationTextRangeFacade | null>
  }
>
type RootPresentation = NonNullable<OfficeSDK['presentation']>

export type EditorFacadeContractAssertions = [
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['title']>['setTitle'],
      (title: string) => Promise<void>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['title']>['addChangedListener'],
      (listener: (title: string) => void) => () => void
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['locks']>['addRangeLock'],
      (options: import('./OfficeSDK').AddRangeLockParams) => Promise<void>
    >
  >,
  Assert<
    IsAssignable<NonNullable<OfficeSDK['history']>['show'], () => Promise<void>>
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['mention']>['locateCellByGuid'],
      (
        guid: string,
        notificationType?: import('./OfficeSDK').MentionTypes
      ) => Promise<void>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['version']>['createRevision'],
      (
        options?: import('./OfficeSDK').RevisionCreateOptions
      ) => Promise<
        undefined | null | import('./types/Document').DocumentErrorMessage
      >
    >
  >,
  Assert<IsAssignable<RootPresentation, BasicPresentationFacade>>,
  Assert<IsAssignable<RootPresentation['start'], () => Promise<void>>>,
  Assert<IsAssignable<RootPresentation['quit'], () => Promise<void>>>,
  Assert<
    IsAssignable<PresentationFacade['start'], (index?: number) => Promise<void>>
  >,
  Assert<
    IsAssignable<PresentationFacade['startFromCurrent'], () => Promise<void>>
  >,
  Assert<
    IsAssignable<PresentationFacade['startRemoteLive'], () => Promise<void>>
  >,
  Assert<
    IsAssignable<PresentationFacade['startSpeakerView'], () => Promise<void>>
  >,
  Assert<
    IsAssignable<
      PresentationFacade['addChangeListener'],
      (listener: (payload: unknown) => void) => () => void
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['batchChanges']>,
      <T>(callback: () => T | Promise<T>) => Promise<Awaited<T>>
    >
  >,
  Assert<IsAssignable<NonNullable<OfficeSDK['print']>, () => Promise<void>>>,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['export']>,
      (type: string) => Promise<void>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['setFocus']>,
      (isFocus: boolean) => Promise<void>
    >
  >,
  Assert<
    IsAssignable<
      DocsSelection['getRange'],
      (value?: DocsRangeValue) => Promise<DocsRangeFacade | null>
    >
  >,
  Assert<
    IsAssignable<
      PresentationSelection['getTextRange'],
      (
        value?: PresentationTextRangeValue
      ) => Promise<PresentationTextRangeFacade | null>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['tables']>['getAll'],
      () => Promise<DocsTableFacade[]>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['workbook']>['getWorksheets'],
      () => Promise<import('./OfficeSDK').SheetWorksheetFacade[]>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['activeSheet']>['getCell'],
      (
        row: number,
        column: number
      ) => Promise<import('./OfficeSDK').SheetCellFacade | null>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['slides']>['getCurrentSlide'],
      () => Promise<PresentationSlideFacade>
    >
  >,
  Assert<
    IsAssignable<
      PresentationSlideFacade['insertTextBox'],
      (
        options: import('./OfficeSDK').PresentationTextBoxOptions
      ) => Promise<PresentationTextShape>
    >
  >,
  Assert<
    IsAssignable<
      PresentationSlideFacade['insertShape'],
      (
        options: Extract<
          import('./OfficeSDK').PresentationInsertShapeOptions,
          { type: import('./OfficeSDK').PresentationLineShapeType }
        >
      ) => Promise<PresentationShapeBase>
    >
  >,
  Assert<
    IsAssignable<
      PresentationSlideFacade['insertShape'],
      (
        options: Extract<
          import('./OfficeSDK').PresentationInsertShapeOptions,
          { type: import('./OfficeSDK').PresentationTextShapeType }
        >
      ) => Promise<PresentationTextShape>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['text']>['apply'],
      (
        format: Partial<EditorTextFormat>,
        range?: PresentationTextRangeValue
      ) => Promise<Partial<EditorTextFormat>>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['activeSheet']>['getSelections'],
      () => Promise<SheetSelection[] | null>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['activeSheet']>['getRange'],
      (value: SheetRangeValue) => Promise<SheetRangeFacade | null>
    >
  >,
  Assert<
    IsEqual<
      SheetRangeFacade['getBounding'],
      () => Promise<{
        left: number
        top: number
        width: number
        height: number
      } | null>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['selections']>['getAll'],
      () => Promise<SheetRangeValue[]>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['charts']>['addChartFromSelection'],
      (
        params?: import('./OfficeSDK').AddChartFromSelectionParams
      ) => Promise<AddChartFromSelectionResult | undefined>
    >
  >,
  Assert<
    IsAssignable<
      PresentationSelection['getSelectedShapes'],
      (ids?: string[]) => Promise<PresentationShape[] | null>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['zoom']>['setPercentage'],
      (percentage: number) => Promise<void>
    >
  >,
  Assert<
    IsAssignable<
      NonNullable<OfficeSDK['eventSubscription']>['addLoadedListener'],
      (listener: () => void) => () => void
    >
  >,
  Assert<
    IsEqual<
      SheetSelection['getRange'],
      (value?: SheetRangeValue) => Promise<SheetRangeFacade | null>
    >
  >,
  Assert<IsEqual<SheetSelection['id'], string>>,
  Assert<
    IsEqual<
      PresentationShapeBase['setFill'],
      (fill: import('./OfficeSDK').PresentationShapeFill) => Promise<void>
    >
  >,
  Assert<IsEqual<PresentationShapeBase['remove'], () => Promise<void>>>,
  Assert<
    IsEqual<
      PresentationTableItem['getCell'],
      (row: number, column: number) => Promise<PresentationTableCell | null>
    >
  >,
  Assert<
    IsEqual<
      PresentationTableItem['getRange'],
      (
        range: PresentationTableSelection
      ) => Promise<PresentationTableRange | null>
    >
  >,
  Assert<
    IsEqual<
      PresentationTableItem['insertRows'],
      (
        index: number,
        count: number,
        placement?: 'before' | 'after'
      ) => Promise<void>
    >
  >,
  Assert<IsEqual<PresentationTableCell['clearStyle'], () => Promise<void>>>,
  Assert<IsEqual<PresentationTableRange['setSpan'], () => Promise<void>>>
]
