Dockerfile, requirements.txt, app.pyをディレクトリに作成します。

csv_filesディレクトリを上記のディレクトリ内に作成し、その中にCSVファイルを配置します。

以下のコマンドを実行して、Dockerイメージをビルドし、コンテナを実行します。


docker build -t python-task .
docker run -v $(pwd):/app python-task


これで、コンテナ内でapp.pyが実行され、指定された処理が実行されて、結果のExcelファイルがproject_directoryにoutput.xlsxとして保存されます。


Dockerを停止するには、
docker stop [コンテナIDまたは名前]
もしくは
docker stop $(docker ps -aq)　→全てのコンテナが対象
でコンテナを停止し、
docker rm [コンテナIDまたは名前]
でコンテナを削除します。

Dockerイメージを削除するには、次のコマンドを使用します
docker rmi [イメージ名]


DockerコンテナのIDを調べるには、
docker ps -a
でリスト表示されるので、わかります。
コンテナ名を知っている場合は、
docker inspect [コンテナ名]
で調べられます。