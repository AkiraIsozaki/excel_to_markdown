# 基本設計書

# 基本設計書（データベース設計）

| システム名 | 在庫管理システム |
| --- | --- |
| 設計フェーズ | 基本設計 |
| 作成者 | 鈴木 花子 |
| 作成日 | 2025-02-01 |

## 1. 設計目的

本書は在庫管理システムのデータベース基本設計を定義する。

各テーブルの論理設計・物理設計・テーブル間の関連を明確にすることを目的とする。

## 2. テーブル一覧

| No. | テーブル名（論理） | テーブル名（物理） | 区分 | 説明 |
| --- | --- | --- | --- | --- |
| 1 | 商品マスタ | M_ITEM | マスタ | 商品の基本情報を管理する |
| 2 | 在庫テーブル | T_STOCK | トラン | 現在の在庫数量を管理する |
| 3 | 入庫履歴 | T_RECEIPT | トラン | 入庫トランザクションを記録する |
| 4 | 出庫履歴 | T_SHIPMENT | トラン | 出庫トランザクションを記録する |
| 5 | 倉庫マスタ | M_WAREHOUSE | マスタ | 倉庫・ロケーション情報を管理する |

## 3. テーブル詳細設計

### 3.1 商品マスタ（M_ITEM）

| カラム名 | データ型 | NULL | PK/FK | 説明 |
| --- | --- | --- | --- | --- |
| item_code | VARCHAR(20) | NOT NULL | PK | 商品コード（JANコード） |
| item_name | VARCHAR(100) | NOT NULL |  | 商品名称 |
| category_id | INTEGER | NOT NULL | FK | カテゴリID（カテゴリマスタ参照） |
| unit_price | DECIMAL(10,2) | NOT NULL |  | 単価 |
| reorder_qty | INTEGER | NOT NULL |  | 発注点数量（アラート閾値） |
| created_at | TIMESTAMP | NOT NULL |  | 登録日時 |
| updated_at | TIMESTAMP | NOT NULL |  | 更新日時 |

※ 全テーブルに created_by（登録者ID）・updated_by（更新者ID）を共通カラムとして追加すること

### 3.2 在庫テーブル（T_STOCK）

| カラム名 | データ型 | NULL | PK/FK | 説明 |
| --- | --- | --- | --- | --- |
| stock_id | BIGINT | NOT NULL | PK | 在庫ID（自動採番） |
| item_code | VARCHAR(20) | NOT NULL | FK | 商品コード |
| warehouse_id | INTEGER | NOT NULL | FK | 倉庫ID |
| quantity | INTEGER | NOT NULL |  | 現在庫数 |
| updated_at | TIMESTAMP | NOT NULL |  | 最終更新日時 |



---

# ER図メモ

## テーブル間リレーション

#### 主要リレーション定義

| M_ITEM（商品マスタ） | T_STOCK（在庫テーブル） | 1 : N | 商品は複数倉庫に在庫を持つ |
| --- | --- | --- | --- |
| M_WAREHOUSE（倉庫マスタ） | T_STOCK（在庫テーブル） | 1 : N | 倉庫は複数商品の在庫を持つ |
| M_ITEM（商品マスタ） | T_RECEIPT（入庫履歴） | 1 : N | 入庫は必ず商品マスタに紐づく |
| M_ITEM（商品マスタ） | T_SHIPMENT（出庫履歴） | 1 : N | 出庫は必ず商品マスタに紐づく |

#### 制約事項

- 在庫数（quantity）は 0 以上の整数であること（CHECK制約）
- 商品マスタを削除する場合は在庫テーブル・履歴テーブルの参照を確認すること（論理削除推奨）
- 在庫テーブルの更新は必ずトランザクション内で行うこと
