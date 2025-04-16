#!/bin/bash

# サーバーのURLを設定
SERVER_URL="http://localhost:8000/rpc"

# リクエストを送信
curl -X POST $SERVER_URL \
  -H "Content-Type: application/json" \
  -d @test_data_request.json

echo ""
echo "完了しました"
