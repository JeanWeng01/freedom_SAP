FROM python:3.13-slim-bookworm

# Pinned Chrome major. Bump after verifying compatibility on Railway.
# Chrome + chromedriver are both fetched from Chrome for Testing for this major,
# so versions stay in lockstep across rebuilds.
ARG CHROME_MAJOR=140

# Tooling + Chrome runtime dependencies (since we no longer install
# google-chrome-stable via apt, its deps aren't pulled automatically).
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget \
    curl \
    unzip \
    ca-certificates \
    tzdata \
    fonts-liberation \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libatspi2.0-0 \
    libcairo2 \
    libcups2 \
    libdbus-1-3 \
    libdrm2 \
    libgbm1 \
    libgtk-3-0 \
    libnspr4 \
    libnss3 \
    libpango-1.0-0 \
    libvulkan1 \
    libx11-6 \
    libxcb1 \
    libxcomposite1 \
    libxdamage1 \
    libxext6 \
    libxfixes3 \
    libxkbcommon0 \
    libxrandr2 \
    libxshmfence1 \
    libxss1 \
    xdg-utils \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install Chrome + matching chromedriver from Chrome for Testing (pinned major).
# Using the same source for both guarantees they match patch-version exactly.
RUN CHROME_VERSION=$(curl -fsSL "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_${CHROME_MAJOR}") \
    && echo "Installing Chrome/Chromedriver ${CHROME_VERSION}" \
    && wget -q -O /tmp/chrome.zip "https://storage.googleapis.com/chrome-for-testing-public/${CHROME_VERSION}/linux64/chrome-linux64.zip" \
    && unzip -q /tmp/chrome.zip -d /opt/ \
    && ln -sf /opt/chrome-linux64/chrome /usr/local/bin/google-chrome \
    && rm /tmp/chrome.zip \
    && wget -q -O /tmp/chromedriver.zip "https://storage.googleapis.com/chrome-for-testing-public/${CHROME_VERSION}/linux64/chromedriver-linux64.zip" \
    && unzip -j /tmp/chromedriver.zip "chromedriver-linux64/chromedriver" -d /usr/local/bin/ \
    && chmod +x /usr/local/bin/chromedriver \
    && rm /tmp/chromedriver.zip \
    && echo "Installed: $(google-chrome --version) | $(chromedriver --version)"

WORKDIR /app

# Copy and install Python deps
COPY sap_bot/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy bot code
COPY sap_bot/ .

# Railway sets PORT env var
ENV HEADLESS=true
ENV PYTHONUNBUFFERED=1
ENV TZ=America/New_York

EXPOSE 8080

CMD ["python", "server.py"]
