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
                    # Siapkan file untuk hasil tes agar dapat ditulis oleh non-root user (appuser) di dalam container
                    touch test-results.xml
                    chmod 666 test-results.xml
                    
                    # Jalankan test di dalam container yang baru dibangun, mount direktori tests dan file output
                    docker run --rm \
                        -v "${WORKSPACE}/tests:/app/tests" \
                        -v "${WORKSPACE}/test-results.xml:/app/test-results.xml" \
                        ${IMAGE_NAME}:${IMAGE_TAG} pytest tests/ -v --junitxml=test-results.xml
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
