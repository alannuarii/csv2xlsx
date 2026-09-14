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

        stage('Prepare Environment & Dependencies') {
            steps {
                sh '''
                    python3 -m venv .venv
                    . .venv/bin/activate
                    pip install --upgrade pip
                    pip install -r requirements.txt
                '''
            }
        }

        stage('Automated Tests') {
            steps {
                sh '''
                    . .venv/bin/activate
                    pytest tests/ -v --junitxml=test-results.xml
                '''
            }
            post {
                always {
                    junit testResults: 'test-results.xml', allowEmptyResults: true
                }
            }
        }

        stage('Build Docker Image') {
            steps {
                sh '''
                    docker build -t ${IMAGE_NAME}:${IMAGE_TAG} -t ${IMAGE_NAME}:latest .
                '''
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
