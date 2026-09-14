pipeline {
    agent any

    environment {
        IMAGE_NAME = 'csv2xlsx'
        IMAGE_TAG = "${env.BUILD_NUMBER ?: 'latest'}"
        HOST_PORT = '3021'
        CONTAINER_PORT = '8000'
    }

    options {
        timeout(time: 15, unit: 'MINUTES')
        disableConcurrentBuilds()
    }

    stages {
        stage('Checkout') {
            steps {
                checkout scm
            }
        }

        stage('Build Docker Image') {
            steps {
                sh '''
                    docker build -t ${IMAGE_NAME}:${IMAGE_TAG} -t ${IMAGE_NAME}:latest .
                '''
            }
        }

        stage('Automated Tests') {
            steps {
                sh '''
                    # 1. Buat container dari image (tanpa start)
                    CONTAINER_ID=$(docker create ${IMAGE_NAME}:${IMAGE_TAG} pytest tests/ -v --junitxml=test-results.xml)
                    
                    # 2. Copy direktori tests dari workspace ke dalam container
                    docker cp tests $CONTAINER_ID:/app/
                    
                    # 3. Jalankan container untuk mengeksekusi pytest dan abaikan error sementara
                    docker start -a $CONTAINER_ID || true
                    
                    # 4. Ambil exit code dari pytest
                    EXIT_CODE=$(docker inspect $CONTAINER_ID --format='{{.State.ExitCode}}')
                    
                    # 5. Copy file hasil test (test-results.xml) dari dalam container ke workspace Jenkins
                    docker cp $CONTAINER_ID:/app/test-results.xml ./test-results.xml || true
                    
                    # 6. Hapus container test
                    docker rm $CONTAINER_ID
                    
                    # 7. Kembalikan exit code pytest agar Jenkins tahu test gagal/sukses
                    exit $EXIT_CODE
                '''
            }
            post {
                always {
                    junit testResults: 'test-results.xml', allowEmptyResults: true
                }
            }
        }

        stage('Deploy Container') {
            steps {
                sh '''
                    docker stop ${IMAGE_NAME} || true
                    docker rm ${IMAGE_NAME} || true
                    docker run -d \
                        --name ${IMAGE_NAME} \
                        --restart always \
                        -p ${HOST_PORT}:${CONTAINER_PORT} \
                        ${IMAGE_NAME}:latest
                '''
            }
        }
    }

    post {
        success {
            echo "Pipeline selesai dengan sukses! Container ${IMAGE_NAME} berjalan pada port ${HOST_PORT}:${CONTAINER_PORT} dengan --restart always."
        }
        failure {
            echo "Pipeline gagal saat proses eksekusi build atau test."
        }
        always {
            cleanWs deleteDirs: true, notFailBuild: true
        }
    }
}
