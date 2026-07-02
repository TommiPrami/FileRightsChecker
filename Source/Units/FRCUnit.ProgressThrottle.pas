unit FRCUnit.ProgressThrottle;

// Bridges TFileRightsChecker progress callbacks (high-frequency) to a UI update
// callback (low-frequency). The checker fires OnTest after every diagnostic call —
// potentially thousands per second on big trees — so updating the form directly
// would pin the message pump. The throttle batches updates to at most one per
// UpdateIntervalMS, but always emits the very first call and the final 100% so
// the operator sees both ends of the run.

interface

uses
  System.Diagnostics, System.SysUtils, FRCUnit.FileRightsChecker;

type
  // Form-side hook that actually puts text on screen.
  TShowProgressEvent = procedure(const AProgressText: string) of object;

  TProgressThrottle = class(TObject)
  strict private
    FStopwatch: TStopwatch;
    FUpdateIntervalMS: Cardinal;
    FOnShowProgress: TShowProgressEvent;
    FHasReportedAny: Boolean;
    FLastShownPercent: Double;
    FCaptionPrefix: string;
    function ShouldShow(const AProgress: Double): Boolean;
    procedure FireShow(const AText: string);
  public
    // ACaptionPrefix is the form's original caption — kept untouched and prepended to
    // every progress update so the caption reads "<original>  - <progress>".
    // When 100% is reached we also fire one final update with just the prefix so the
    // caption is restored to its original state.
    constructor Create(const ACaptionPrefix: string; const AUpdateIntervalMS: Cardinal = 100);

    // Resets timing and "have we shown anything yet" tracking — call between runs.
    procedure Reset;

    // Hooked to TFileRightsChecker.OnTest.
    procedure HandleTest(const AType: TFileSystemType; const AName: string;
      const ATestCount: Integer; const AErrorsCount: Integer; const AProgress: Double);

    property UpdateIntervalMS: Cardinal read FUpdateIntervalMS write FUpdateIntervalMS;
    property OnShowProgress: TShowProgressEvent read FOnShowProgress write FOnShowProgress;
    property CaptionPrefix: string read FCaptionPrefix;
  end;

implementation

uses
  System.Math;

const
  // Used by SameValue when comparing percentages. 0.001 is well below any
  // sub-percent resolution we'd surface to a human reader.
  PROGRESS_EPSILON = 0.001;

constructor TProgressThrottle.Create(const ACaptionPrefix: string; const AUpdateIntervalMS: Cardinal = 100);
begin
  inherited Create;

  FCaptionPrefix := ACaptionPrefix;
  FUpdateIntervalMS := AUpdateIntervalMS;
  FStopwatch := TStopwatch.StartNew;
  FHasReportedAny := False;
  FLastShownPercent := -1;
end;

procedure TProgressThrottle.Reset;
begin
  FStopwatch := TStopwatch.StartNew;
  FHasReportedAny := False;
  FLastShownPercent := -1;
end;

function TProgressThrottle.ShouldShow(const AProgress: Double): Boolean;
begin
  // First update of a run — operator wants to see "we started".
  if not FHasReportedAny then
    Exit(True);

  // Always show 0% explicitly (covered by the first-call branch too, but cheap to be
  // explicit about it).
  if SameValue(AProgress, 0.0, PROGRESS_EPSILON) then
    Exit(True);

  // Always show 100% — without this the final tail of work could be silenced by the
  // throttle window and the operator would think it's hung.
  if SameValue(AProgress, 100.0, PROGRESS_EPSILON) then
    Exit(True);

  // Skip identical percentages so we don't redraw the same string repeatedly.
  if SameValue(AProgress, FLastShownPercent, PROGRESS_EPSILON) then
    Exit(False);

  Result := FStopwatch.ElapsedMilliseconds >= FUpdateIntervalMS;
end;

procedure TProgressThrottle.FireShow(const AText: string);
begin
  if Assigned(FOnShowProgress) then
    FOnShowProgress(AText);

  // Restart the throttle window from "now", so the next update is at least
  // UpdateIntervalMS away.
  FStopwatch := TStopwatch.StartNew;
  FHasReportedAny := True;
end;

procedure TProgressThrottle.HandleTest(const AType: TFileSystemType; const AName: string; const ATestCount: Integer;
  const AErrorsCount: Integer; const AProgress: Double);
var
  LItemLabel: string;
  LProgressText: string;
  LIsFinal: Boolean;
begin
  LIsFinal := SameValue(AProgress, 100.0, PROGRESS_EPSILON);

  if not ShouldShow(AProgress) then
    Exit;

  if AType = fstDirectory then
    LItemLabel := 'directory'
  else
    LItemLabel := 'file';

  // FormatFloat respects the current locale's ThousandSeparator and DecimalSeparator,
  // so on a Finnish locale this produces "1 666" and "66,60" as the user expects.
  // Format with %f would NOT respect locale by default.
  LProgressText := Format('Current %s: %s - Tests run so far: %s, Errors: %d - %s%%',
    [LItemLabel, AName, FormatFloat('#,##0', ATestCount), AErrorsCount,
     FormatFloat('0.00', AProgress)]);

  // Prefix the saved caption so the host form keeps its identity while we update.
  FireShow(FCaptionPrefix + '  - ' + LProgressText);
  FLastShownPercent := AProgress;

  // On 100% restore the caption to its original text. Deliberately NOT one-shot:
  // with epsilon comparison, near-final updates (99.999%+ on very large scans) also
  // count as "final", and a one-shot guard would let a later progress write win and
  // leave the caption stuck. Restoring repeatedly is idempotent — same text.
  if LIsFinal then
    FireShow(FCaptionPrefix);
end;

end.
