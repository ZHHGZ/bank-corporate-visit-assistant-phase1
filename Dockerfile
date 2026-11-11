# 企业报告聚合 API 生产镜像
FROM python:3.11-slim

# UTF-8 语言环境，确保中文企业名/文件名正常写入
ENV LANG=C.UTF-8 \
    LC_ALL=C.UTF-8 \
    PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    TZ=Asia/Shanghai

WORKDIR /app

# 先装依赖以利用镜像层缓存
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

# 拷贝应用代码与正式报告所需的配置/资源
COPY app ./app
COPY configs ./configs
COPY develop_docs ./develop_docs

# 报告输出与原始响应上下文缓存目录（可挂载卷持久化）
RUN mkdir -p temp \
    && useradd --create-home --uid 10001 appuser \
    && chown -R appuser:appuser /app
USER appuser

EXPOSE 8001

# 健康检查（slim 镜像无 curl，用 python 探活）
HEALTHCHECK --interval=30s --timeout=5s --start-period=15s --retries=3 \
    CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8001/health', timeout=3).status==200 else 1)"

# 生产模式启动（不带 --reload）
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8001"]
