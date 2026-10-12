@Library('pipeline-library') _

// Pipeline for ClickAI Product Normalizer - PROD environment
buildAndDeploy(
    // === REQUIRED PARAMETERS ===
    environment: 'prod',
    skipConfirmation: true,
    appName: 'product-normalizer',
    repo: 'https://github.com/nvtthanh11/clickai-product-normalizer.git',
    branch: 'main',
    credentialsId: 'github-pat',
    registry: 'k8s.io.reg.vnetwork.dev/clickai',
    kubeconfigPath: '/var/jenkins_home/.kube/config',
    dockerCredentialsId: 'docker-registry-credentials',

    // === OPTIONAL PARAMETERS ===
    // Docker build configuration
    dockerfile: './Dockerfile',
    buildContext: '.',
    platform: 'linux/amd64',

    // Kubernetes deployment configuration (PROD)
    deployment: 'clickai-prod-product-normalizer',
    namespace: 'prod',
    kubeContext: 'prod',
    containerName: 'product-normalizer',

    // Deployment options
    setImage: true,
    waitForRollout: true,
    rolloutTimeout: '5m',

    // Cleanup
    cleanupImages: true
)
