"""
AI 문서분류 훈련 스크립트
ETA LIMS — 분석항목 파서 자동 분류 모델 (TF-IDF + LogisticRegression → ONNX)

사용법:
    python Scripts/ai_train.py

입력:  Data/ai_training_data.json
출력:  Data/분석항목_분류.onnx
"""

import json
import sys
import os
import numpy as np
from sklearn.pipeline import Pipeline
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import cross_val_score
from skl2onnx import convert_sklearn
from skl2onnx.common.data_types import StringTensorType

# ── 경로 설정 ──────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_FILE  = os.path.join(BASE_DIR, "Data", "ai_training_data.json")
ONNX_FILE  = os.path.join(BASE_DIR, "Data", "분석항목_분류.onnx")


def load_training_data():
    with open(DATA_FILE, encoding="utf-8") as f:
        data = json.load(f)

    samples = data.get("samples", [])
    if not samples:
        print("[오류] 훈련 데이터가 없습니다. 파일을 먼저 추가해주세요.")
        sys.exit(1)

    texts  = [s["text"]  for s in samples]
    labels = [s["label"] for s in samples]

    label_counts = {}
    for l in labels:
        label_counts[l] = label_counts.get(l, 0) + 1

    print(f"[데이터] 총 {len(texts)}건")
    for label, cnt in sorted(label_counts.items()):
        print(f"  {label}: {cnt}건")

    return texts, labels


def train(texts, labels):
    pipeline = Pipeline([
        ("tfidf", TfidfVectorizer(
            max_features  = 3000,
            ngram_range   = (1, 2),
            sublinear_tf  = True,
            min_df        = 1,
        )),
        ("clf", LogisticRegression(
            max_iter     = 1000,
            C            = 1.0,
            solver       = "lbfgs",
        )),
    ])

    # 클래스가 2종 이상이고 샘플이 충분할 때 교차검증
    unique_labels = list(set(labels))
    if len(unique_labels) >= 2 and len(texts) >= len(unique_labels) * 2:
        cv = min(5, len(texts) // len(unique_labels))
        if cv >= 2:
            scores = cross_val_score(pipeline, texts, labels, cv=cv, scoring="accuracy")
            print(f"[교차검증] {cv}-fold 정확도: {scores.mean():.1%} (±{scores.std():.1%})")

    pipeline.fit(texts, labels)
    print(f"[훈련 완료] 클래스: {sorted(set(labels))}")
    return pipeline


def export_onnx(pipeline):
    initial_type = [("string_input", StringTensorType([None, 1]))]
    onnx_model = convert_sklearn(
        pipeline,
        initial_types = initial_type,
        target_opset  = 17,
    )

    os.makedirs(os.path.dirname(ONNX_FILE), exist_ok=True)
    with open(ONNX_FILE, "wb") as f:
        f.write(onnx_model.SerializeToString())

    size_kb = os.path.getsize(ONNX_FILE) / 1024
    print(f"[ONNX 저장] {ONNX_FILE} ({size_kb:.1f} KB)")


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
    print("=" * 50)
    print(" ETA AI 문서분류 - 모델 훈련")
    print("=" * 50)

    if not os.path.exists(DATA_FILE):
        print(f"[오류] 훈련 데이터 파일 없음: {DATA_FILE}")
        sys.exit(1)

    texts, labels = load_training_data()
    pipeline = train(texts, labels)
    export_onnx(pipeline)

    print("=" * 50)
    print(" 학습 완료! 분석항목_분류.onnx 생성됨")
    print("=" * 50)


if __name__ == "__main__":
    main()
