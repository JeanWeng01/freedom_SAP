FROM python:3.13-slim

# Install Chrome + dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget \
    gnupg2 \
    unzip \
    curl \
    && wget -q -O - https://dl.google.com/linux/linux_signing_key.pub | gpg --dearmor -o /usr/share/keyrings/google-chrome.gpg \
    && echo "deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome.gpg] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update \
    && apt-get install -y --no-install-recommends google-chrome-stable \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Install chromedriver matching the installed Chrome version (Chrome for Testing).
# Baked into the image so Selenium Manager never has to download at runtime.
RUN CHROME_MAJOR=$(google-chrome --version | awk '{print $3}' | cut -d. -f1) \
    && DRIVER_VERSION=$(curl -fsSL "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_${CHROME_MAJOR}") \
    && wget -q -O /tmp/chromedriver.zip "https://storage.googleapis.com/chrome-for-testing-public/${DRIVER_VERSION}/linux64/chromedriver-linux64.zip" \
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

# Install timezone data for ET
RUN apt-get update && apt-get install -y --no-install-recommends tzdata && rm -rf /var/lib/apt/lists/*

# Railway sets PORT env var
ENV HEADLESS=true
ENV PYTHONUNBUFFERED=1
ENV TZ=America/New_York

EXPOSE 8080

CMD ["python", "server.py"]
