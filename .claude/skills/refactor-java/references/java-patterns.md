# Java リファクタリング パターン集

## 関心の分離とパッケージ可視性

### パターン: パッケージプライベートを優先する

Javaでは、アクセス修飾子なし（パッケージプライベート）が「このクラスはこのパッケージの内部実装」という
意思表示になります。パッケージ外に公開する必要がないものは積極的にパッケージプライベートにします。

```java
// ❌ Anti-pattern: 外部から使われないのに public
package com.example.core;
public class InternalMatrixHelper {
    public static Fraction[] extractRow(Matrix m, int row) { ... }
}

// ✅ Pattern: パッケージプライベートで内部実装を隠蔽
package com.example.core;
class InternalMatrixHelper {
    static Fraction[] extractRow(Matrix m, int row) { ... }
}
```

### パターン: 凝集すべきクラスを同じパッケージに集める

関連するクラス群が複数パッケージに散在していると、変更時の影響範囲が広がります。
「一緒に変わるものは一緒に置く」原則に従います。

```
// ❌ Anti-pattern: EigenvalueResultがcoreではなくdtoパッケージに
com.example.core.EigenvalueCalculator
com.example.dto.EigenvalueResult        ← 別パッケージに分離

// ✅ Pattern: 計算と結果型を同じパッケージに
com.example.core.EigenvalueCalculator
com.example.core.EigenvalueResult       ← 計算と結果は一体
```

### パターン: パッケージ間のファサード

どうしても外部に公開が必要な場合は、インターフェースをパッケージの「窓口」にします。

```java
// パッケージ外に公開するインターフェース（public）
public interface Calculator<T> {
    T calculate(Matrix matrix);
}

// 実装はパッケージプライベート
class EigenvalueCalculatorImpl implements Calculator<EigenvalueResult> {
    @Override
    public EigenvalueResult calculate(Matrix matrix) { ... }
}
```

---

## 完全コンストラクタと値オブジェクト

### パターン: 完全コンストラクタ

オブジェクトの生成時点で、有効な状態を完全に確定させます。
「半分だけ初期化されたオブジェクト」は存在を許しません。

```java
// ❌ Anti-pattern: 段階的な初期化（不完全な状態が存在できる）
class SolverResult {
    private SolutionType type;
    private Fraction[] solution;

    public void setType(SolutionType type) { this.type = type; }
    public void setSolution(Fraction[] solution) { this.solution = solution; }
}

// ✅ Pattern: 完全コンストラクタ（生成直後から有効）
final class SolverResult {
    private final SolutionType type;
    private final Fraction[] solution;

    private SolverResult(SolutionType type, Fraction[] solution) {
        this.type = Objects.requireNonNull(type);
        this.solution = solution == null ? null : solution.clone();
    }

    // 名前付きファクトリメソッドで意図を明確に
    static SolverResult unique(Fraction[] solution) {
        return new SolverResult(SolutionType.UNIQUE, Objects.requireNonNull(solution));
    }

    static SolverResult noSolution() {
        return new SolverResult(SolutionType.NO_SOLUTION, null);
    }
}
```

### パターン: 不変性（Immutability）

フィールドを `final` にし、コレクションは防御的コピーをとります。

```java
// ❌ Anti-pattern: 可変コレクションをそのまま保持
class EigenvalueResult {
    private List<Fraction> eigenvalues;

    public EigenvalueResult(List<Fraction> eigenvalues) {
        this.eigenvalues = eigenvalues; // 呼び出し元が変更できる
    }
}

// ✅ Pattern: 不変コピーを保持
final class EigenvalueResult {
    private final List<Fraction> eigenvalues;

    EigenvalueResult(List<Fraction> eigenvalues) {
        this.eigenvalues = List.copyOf(eigenvalues); // 不変コピー
    }

    List<Fraction> eigenvalues() {
        return eigenvalues; // List.copyOfは既にimmutable
    }
}
```

### アンチパターン: 不要なstatic

staticメソッドは「クラスのインスタンス不要」という強い意思表示ですが、
乱用するとテストが難しくなり、依存関係が隠れます。

```java
// ❌ Anti-pattern: 何でもstaticにする
class MatrixUtils {
    public static Matrix add(Matrix a, Matrix b) { ... }
    public static Matrix multiply(Matrix a, Matrix b) { ... }
    // staticメソッドはモックできない
}

// ✅ Pattern: インスタンスメソッド（またはMatrixクラス自身のメソッド）
final class Matrix {
    Matrix add(Matrix other) { ... }      // 自然なオブジェクトの振る舞い
    Matrix multiply(Matrix other) { ... }
}
```

**staticが適切なケース**:
- `Math.abs()` のような純粋な数学関数
- `List.of()` のような静的ファクトリメソッド
- 定数（`static final`）

---

## プリミティブ執着の回避

### パターン: Value Object（値オブジェクト）の導入

「ドメインの概念」を持つプリミティブは Value Object に包みます。
Value Object は：等値比較（equals/hashCode）を値で行い、不変で、自己検証します。

**型の選択指針**:
- **Java 16+ `record` を第一候補にする** — `equals`/`hashCode`/`toString` が自動生成され、コードが簡潔
- `final class` を選ぶのは: パッケージプライベートにしたい場合、または `record` では記述できない複雑な検証が必要な場合

```java
// ❌ Anti-pattern: int で行・列を表現
void highlightCell(int row, int col) {
    // row と col を逆に渡してもコンパイルエラーにならない
}

// ✅ Pattern 1: record（推奨、Java 16+）
// equals/hashCode/toString が自動生成される
record CellPosition(int row, int col) {
    // コンパクトコンストラクタで検証
    CellPosition {
        if (row < 0 || col < 0) throw new IllegalArgumentException(
            "row and col must be non-negative: row=" + row + ", col=" + col);
    }
}

// ✅ Pattern 2: final class（パッケージプライベートにしたい場合等）
final class CellPosition {
    private final int row;
    private final int col;

    CellPosition(int row, int col) {
        if (row < 0 || col < 0) throw new IllegalArgumentException("...");
        this.row = row;
        this.col = col;
    }

    int row() { return row; }
    int col() { return col; }

    @Override
    public boolean equals(Object o) {
        if (!(o instanceof CellPosition other)) return false;
        return row == other.row && col == other.col;
    }

    @Override
    public int hashCode() { return Objects.hash(row, col); }
}

void highlightCell(CellPosition position) {
    // 間違った順序を渡せない
}
```

**⚠️ テスト内部クラス名の衝突に注意**:
テストファイル内に同名の `@Nested` クラスや内部クラスがある場合、コンパイルエラーになります。

```java
// テスト内部クラスが Value Object と同名だと衝突
class StepViewModelTest {
    // ❌ PivotCell を追加したとき、これと衝突する
    @Nested class PivotCell {  // → PivotCellProperty などにリネームが必要
        @Test void ...
    }
}
```

Value Object を追加する前に `Grep` で対象クラス名がテストに存在しないか確認してください。

### パターン: 区分をenumで表現

状態・種類・カテゴリを文字列定数やintで表現している場合、enumに替えます。

```java
// ❌ Anti-pattern: String で状態を表現
String status = "ACTIVE";
if (status.equals("ACTIVE")) { ... }
// タイポしてもコンパイルエラーにならない

// ✅ Pattern: enum で表現
enum AccountStatus {
    ACTIVE, SUSPENDED, CLOSED;
}
AccountStatus status = AccountStatus.ACTIVE;
if (status == AccountStatus.ACTIVE) { ... }
```

### パターン: 計算結果型を作る

メソッドが複数の値を返したい場合に `int[]` や `Object[]` を使うのは
プリミティブ執着です。専用の結果型を定義します。

```java
// ❌ Anti-pattern: int配列で「ピボット位置」を返す
int[] findPivot(Matrix m) {
    return new int[]{row, col}; // [0]が行、[1]が列... ドキュメントなしでは不明
}

// ✅ Pattern: record を使った専用の結果型（推奨）
record PivotPosition(int row, int col) {}

Optional<PivotPosition> findPivot(Matrix m) { ... }
```

---

## 適切な名前設計

### パターン: 名前がビジネスロジックを表現する

```java
// ❌ Anti-pattern: 実装詳細が名前に漏れている
class ArrayBasedMatrix { }         // 内部実装を名前に出してはいけない
void processData(List<int[]> data) { } // "data"は何も語らない

// ✅ Pattern: 概念・役割で命名
class Matrix { }                   // 数学的概念
void applyRowOperations(List<RowOperationStep> steps) { }
```

### パターン: メソッド名は「何をするか」を表す

```java
// ❌ Anti-pattern: 実装手順を説明する名前
void iterateAndSumElements() { }

// ✅ Pattern: 意図（目的）を説明する名前
Fraction computeTrace() { }   // 対角成分の和という数学的概念
```

### パターン: boolean 返却メソッド

```java
// ❌ Anti-pattern: 肯定か否定か不明
boolean check() { }
boolean process() { }

// ✅ Pattern: is/has/can/should で始める
boolean isSquare() { }           // 正方行列かどうか
boolean hasSolution() { }        // 解を持つかどうか
boolean canDiagonalize() { }     // 対角化可能かどうか
```

---

## 適切なインターフェースの利用

### パターン: 依存は抽象（インターフェース）に向ける

テスト時にモックしたい処理、または将来的に実装を差し替える可能性がある処理は
インターフェースで抽象化します。

```java
// ❌ Anti-pattern: 具象クラスに直接依存
class MatrixService {
    private final JsonStorage storage; // Jsonに依存してしまう

    void save(Matrix m, File f) throws IOException {
        storage.save(m, f);
    }
}

// ✅ Pattern: インターフェースに依存
interface MatrixStorage {
    void save(List<Matrix> matrices, List<String> names, File file) throws IOException;
    MatrixFileData load(File file) throws IOException;
}

class MatrixService {
    private final MatrixStorage storage; // 実装に依存しない

    MatrixService(MatrixStorage storage) {
        this.storage = Objects.requireNonNull(storage);
    }
}

// JsonStorageは実装の一つ
class JsonStorage implements MatrixStorage { ... }
```

### アンチパターン: 過度な抽象化

インターフェースの導入は慎重に。実装が1つしかなく、テストでのモック化も不要なら
インターフェースは不要です。

```java
// ❌ Anti-pattern: 無意味なインターフェース（実装が1つで差し替え不要）
interface MatrixAdder {
    Matrix add(Matrix a, Matrix b);
}
class MatrixAdderImpl implements MatrixAdder { ... }

// ✅ Pattern: Matrix自身のメソッドで十分
final class Matrix {
    Matrix add(Matrix other) { ... }
}
```

**インターフェース導入の判断基準**:
- テストでモックしたい（外部I/O、時間依存、ランダム性）
- 実装を差し替える実際の要件がある（設定ファイル・DBによる切り替えなど）
- プラグインポイントとして明示的に設計している

---

---

## 重複コードの抽出

### パターン: 繰り返しの try-catch をヘルパーメソッドに集約

同じ「null チェック + try-catch + エラーダイアログ」が複数のイベントハンドラーに存在するのは
重複コードです。共通の処理を名前付きメソッドに抽出します。

```java
// ❌ Anti-pattern: 同じ構造が6か所に繰り返される
@FXML void handleInverse() {
    final Matrix matrix = getSelectedOrCurrentMatrix();
    if (matrix == null) { return; }
    try {
        final InverseResult result = MatrixService.calculateInverse(matrix);
        // ... 処理
    } catch (final IllegalArgumentException e) {
        DialogHelper.showError("逆行列エラー", e.getMessage());
    }
}

@FXML void handleDeterminant() {
    final Matrix matrix = getSelectedOrCurrentMatrix();
    if (matrix == null) { return; }     // ← 同じ構造
    try {
        final DeterminantResult result = MatrixService.calculateDeterminant(matrix);
        // ... 処理
    } catch (final IllegalArgumentException e) {
        DialogHelper.showError("行列式エラー", e.getMessage());
    }
}

// ✅ Pattern: 共通のテンプレートを抽出
// 「行列を取得して処理する」という構造に名前を付ける
private void withSelectedMatrix(String operationName, Consumer<Matrix> operation) {
    final Matrix matrix = getSelectedOrCurrentMatrix();
    if (matrix == null) { return; }
    try {
        operation.accept(matrix);
    } catch (final IllegalArgumentException e) {
        DialogHelper.showError(operationName + "エラー", e.getMessage());
    }
}

@FXML void handleInverse() {
    withSelectedMatrix("逆行列", matrix -> {
        final InverseResult result = MatrixService.calculateInverse(matrix);
        // ... 処理
    });
}
```

### パターン: Rule of Three (3度目に抽出)

同じコードブロックが3箇所以上出現したら、抽出を検討する目安です。
1〜2箇所なら許容範囲であることが多いです。

```java
// 1回目: そのまま書く
// 2回目: 「もしかして重複?」と気づく
// 3回目: 抽出する
```

### アンチパターン: 過剰なDRY化

一度しか使われない処理を「再利用性のために」抽出するのは過剰です。
抽象化のコストを上回るメリットがなければ、コードの重複より複雑な構造の方が害になります。

```java
// ❌ やりすぎ: 一度しか使われない処理に抽象化を持ち込む
private <T> T executeMatrixOperation(
    Matrix matrix, String errorTitle,
    Supplier<T> operation, Predicate<T> successCheck,
    Consumer<T> onSuccess) { ... }
```

---

## メソッドの長さと複雑度

### パターン: 長いメソッドを意図を表す名前の小さなメソッドに分割

長いメソッドの問題は「長さ」ではなく「複数のことをしている」点です。
それぞれの処理に意図を表す名前を付けることで、読み手が理解しやすくなります。

```java
// ❌ Anti-pattern: 40行の handleVisualize2D — 検証・計算・描画・ウィンドウ表示が混在
@FXML void handleVisualize2D() {
    final Matrix matrix = getSelectedOrCurrentMatrix();
    if (matrix == null) { ... }
    if (matrix.getRows() != 2 || matrix.getCols() != 2) { ... }
    eigenvalueViewModel.visualize2D(matrix);
    if (eigenvalueViewModel.errorMessageProperty().get() != null) { ... }
    final List<double[]> unitSquare = eigenvalueViewModel.transformedUnitSquareProperty().get();
    // ... さらに20行続く
}

// ✅ Pattern: 責務ごとに private メソッドへ分割
@FXML void handleVisualize2D() {
    final Matrix matrix = getSelectedOrCurrentMatrix();
    if (matrix == null) { return; }
    if (!is2By2Matrix(matrix)) { return; }
    eigenvalueViewModel.visualize2D(matrix);
    if (hasError(eigenvalueViewModel)) { return; }
    openVisualizationWindow();
}

private boolean is2By2Matrix(final Matrix matrix) {
    if (matrix.getRows() != 2 || matrix.getCols() != 2) {
        DialogHelper.showError("2D可視化エラー", "2D可視化は2x2行列でのみ可能です。");
        return false;
    }
    return true;
}

private void openVisualizationWindow() {
    // ウィンドウ生成ロジック
}
```

### パターン: 早期 return で深いネストを解消

複数の条件チェックを早期 return に置き換えると、メソッドの「本体」が浮き彫りになります。

```java
// ❌ Anti-pattern: ネストが深くなる
void process(Matrix matrix) {
    if (matrix != null) {
        if (matrix.isSquare()) {
            if (!matrix.isZero()) {
                // 本体の処理
            }
        }
    }
}

// ✅ Pattern: 早期 return でフラット化
void process(final Matrix matrix) {
    if (matrix == null) { return; }
    if (!matrix.isSquare()) { return; }
    if (matrix.isZero()) { return; }
    // 本体の処理 — ここに到達する条件が明確
}
```

---

## 例外処理の設計

### パターン: catch する例外型を必要最小限に絞る

```java
// ❌ Anti-pattern: Exception を catch — NullPointerException も ArithmeticException も全部飲み込む
try {
    final List<RowOperationStep> steps = MatrixService.performRowReduction(matrix);
    showSteps(steps, "掃き出し法");
} catch (final Exception e) {    // ← 広すぎる
    DialogHelper.showError("掃き出し法エラー", e.getMessage());
}

// ✅ Pattern: スローされる具体的な例外型を catch
try {
    final List<RowOperationStep> steps = MatrixService.performRowReduction(matrix);
    showSteps(steps, "掃き出し法");
} catch (final IllegalArgumentException e) {
    DialogHelper.showError("掃き出し法エラー", e.getMessage());
}
// NullPointerExceptionなど予期しない例外はここで飲み込まず上位に伝播させる
```

### パターン: null 返却を Optional に置き換える

null を返すメソッドは、呼び出し側に null チェックを強制します。
`Optional<T>` を使うと「値がない可能性がある」という契約が型に現れます。

```java
// ❌ Anti-pattern: null を返す — 呼び出し側が必ず null チェックを書く必要がある
private Matrix getSelectedOrCurrentMatrix() {
    final int idx = matrixInputController.getSelectedIndex();
    return idx >= 0 ? matrixViewModel.getStoredMatrices().get(idx)
                    : matrixViewModel.currentMatrixProperty().get();
    // currentMatrixProperty().get() が null を返す可能性がある
}

// ✅ Pattern: Optional で「値がない場合」を型として表現
private Optional<Matrix> getSelectedOrCurrentMatrix() {
    final int idx = matrixInputController.getSelectedIndex();
    if (idx >= 0) {
        return Optional.of(matrixViewModel.getStoredMatrices().get(idx));
    }
    return Optional.ofNullable(matrixViewModel.currentMatrixProperty().get());
}

// 呼び出し側: ifPresent / orElse / map などで自然に処理できる
getSelectedOrCurrentMatrix().ifPresent(matrix -> {
    // nullチェックが不要になる
});
```

**注意**: `Optional` は戻り値型に使うのが適切。フィールド型や引数型への使用は一般的に避ける。

### パターン: 同類の操作で一貫した例外ポリシーを持つ

```java
// ❌ Anti-pattern: メソッドごとに catch する例外型がバラバラ
@FXML void handleInverse()    { ... catch (IllegalArgumentException e) { ... } }
@FXML void handleDeterminant(){ ... catch (IllegalArgumentException e) { ... } }
@FXML void handleRowReduction(){ ... catch (Exception e) { ... } }  // ← 一貫性がない

// ✅ Pattern: 同じカテゴリの操作は同じ例外ポリシーで統一
// — 全ての行列演算ハンドラーは IllegalArgumentException を catch する
// — 予期しない例外は全て上位に伝播させる
```

---

## よくある「やりすぎ」パターン

リファクタリングで陥りがちな過剰設計です。避けてください。

### 不必要なラッパークラス

```java
// ❌ やりすぎ: Listを単純にラップするだけのクラス
class MatrixList {
    private final List<Matrix> matrices;
    MatrixList(List<Matrix> matrices) { this.matrices = List.copyOf(matrices); }
    List<Matrix> getAll() { return matrices; }
    // ビジネスロジックがない。List<Matrix>で十分。
}
```

### 全てをインターフェース化する

```java
// ❌ やりすぎ: 変更されないロジックまでインターフェース化
interface Fraction { ... }
interface Matrix { ... }
// これらは数学的に定義が固定されており、実装が変わる理由がない
```

### 全パッケージプライベート化

```java
// ❌ やりすぎ: 公開APIまでパッケージプライベートにする
// ViewModel層のクラスをViewから使えなくなってしまう
class StepViewModel { ... }  // Viewのpackageからアクセス不可
```
