#!/bin/bash

# 환경 변수 설정 (기본값: prod)
ENV=${1:-prod}
echo "Running with environment: $ENV"

# 스크립트가 위치한 디렉토리의 절대 경로 추출
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# 프로젝트 루트 디렉토리 (scripts의 부모 디렉토리)
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"

echo "Project root: $PROJECT_ROOT"

# Python 가상환경 활성화
source "$PROJECT_ROOT/.venv/bin/activate"

# PYTHONPATH 설정
export PYTHONPATH="$PROJECT_ROOT:$PYTHONPATH"

# 현재 디렉토리 확인
pwd

# main.py 실행
cd "$PROJECT_ROOT/src"
python3 main.py --env "$ENV"
