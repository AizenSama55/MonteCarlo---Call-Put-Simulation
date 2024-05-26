#include <curand_kernel.h>

extern "C" __global__ void MonteCarloKernel(double currentPrice, double meanReturn, double volatility, int days, double* finalPrices)
{
    int idx = blockIdx.x * blockDim.x + threadIdx.x;
    double price = currentPrice;
    double randomShock;

    // Initialize the random state
    curandState state;
    curand_init(1234, idx, 0, &state);

    for (int t = 1; t < days; t++)
    {
        randomShock = meanReturn + curand_normal(&state) * volatility;
        price = price * (1 + randomShock);
    }
    finalPrices[idx] = price;
}
