#include <stdio.h>
#include <math.h>
#include <time.h>
#include <stdlib.h>
#include <string.h>
#include <omp.h>

#define max(a,b) (((a) > (b)) ? (a) : (b))

#define N 256
#define EPS 0.001
#define OMEGA 1.99

double getXY(int i)
{
    return 1.0 * i / (N + 1);
}

int main(int argc, char *argv[])
{
    static double u[N + 2][N + 2];
    static double z[N + 2][N + 2];
    static double TR[N + 2][N + 2];
    static double FUN[N + 2][N + 2];
    int i, j;

    // Prepare arrays 
    for(i = 0; i <= N + 1; i++) {
        for (j = 0; j <= N + 1; j++) {
            double x = getXY(i);
            double y = getXY(j);
            TR[i][j] = (x * x - x + 1) * (y * y - y + 1);
            FUN[i][j] = 4 + 2 * x * x - 2 * x + 2 * y * y - 2 * y;
            if (i == 0 || i == N + 1 || j == 0 || j == N + 1) {
                if (y == 0 || y == 1) {
                    u[i][j] = x * x - x + 1;
                }
                if (x == 0 || x == 1) {
                    u[i][j] = y * y - y + 1;
                }
            }
        }
    }

    memcpy(z, u, sizeof(u));

    FILE *fp;
    fp = fopen("./tmp/relax_iter.txt", "w+");

    #pragma omp parallel
    {
        #pragma omp single
        {
            int numThreads = omp_get_num_threads();
            printf("Threads count is: %d\n", numThreads);
        }
    }

    int k = 0;
    double dmax, dm, d;
    double h = 1.0 / (N + 1);

    double time = omp_get_wtime();
    do
    {
        dmax = 0;
        #pragma omp parallel for shared(z, dmax) private(j, d, dm)
        for (i = 1; i <= N; i++)
        {
            for (j = 1; j <= N; j++)
            {
                z[i][j] = (OMEGA / 4) * (z[i - 1][j] + u[i + 1][j] + z[i][j - 1] + u[i][j + 1] - h * h * FUN[i][j]) + (1 - OMEGA) * u[i][j];
                d = fabs(TR[i][j] - z[i][j]);
                dm = max(dm, d);
            }
            dmax = max(dmax, dm);
        }
        memcpy(u, z, sizeof(z));
        //fprintf(fp, "%d %f\n", k, dmax);
        k++;
    } while (dmax > EPS);
    time = omp_get_wtime() - time;

    printf("TIME SPENT for eps = %f and N = %d is %f s. Number of steps is %d \n", EPS, N, time, k);

    fclose(fp);
    fp = fopen("./tmp/relax.dat", "w+");

    for (i = 0; i < N; i++)
    {
        for (j = 0; j < N; j++)
        {
            double x = getXY(i);
            double y = getXY(j);
            fprintf(fp, "%f %f %f\n", x, y, z[i][j]);
        }
    }

    fclose(fp);
}