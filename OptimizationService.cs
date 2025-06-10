using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Optimizations
{
    public class OptimizationService
    {
        internal class BlockVariant
        {
            public Dictionary<string, double> Values = new Dictionary<string, double>();
            public double[] SumsBySemester = Array.Empty<double>(); // Based on SignificanceCoefficient, used for block-level pruning
        }

        public static Task<List<Dictionary<string, double>>> GenerateValidVariantsAsync(
            List<Discipline> disciplines,
            HashSet<string> excludedDisciplineNames,
            List<(int Sem1, int Sem2, double TargetSum)> semesterPairs)
        {
            return Task.Run(() =>
            {
                var blockVariants = new List<List<BlockVariant>>(semesterPairs.Count);
                var blockVariantsLock = new object();
                int maxVariantsPerBlock = 10; // Keep block-level pruning

                Parallel.ForEach(semesterPairs, pair =>
                {
                    var (sem1, sem2, targetSum) = pair;
                    var blockDisciplines = disciplines.Where(d => d.Semester == sem1 || d.Semester == sem2).ToList();
                    // ... existing code for fixedBlock, variableBlock, baseVariant, valueOptions ...
                    var fixedBlock = blockDisciplines.Where(d => Math.Abs(d.MinWorkload - d.MaxWorkload) < 0.001).ToList();
                    var variableBlock = blockDisciplines.Where(d => Math.Abs(d.MinWorkload - d.MaxWorkload) > 0.001).ToList();
                    
                    var baseVariant = fixedBlock.ToDictionary(d => d.UniqueName, d => d.MinWorkload);
                    var valueOptions = variableBlock
                        .Select(d => GeneratePossibleValues(d.MinWorkload, d.MaxWorkload))
                        .ToList();


                    var localVariants = new List<BlockVariant>();
                    int rejectedCount = 0;

                    void Recurse(int idx, Dictionary<string, double> current)
                    {
                        if (idx == variableBlock.Count)
                        {
                            double sumSem1 = 0, sumSem2 = 0, sumPair = 0, excludedSum = 0;
                            foreach (var d in blockDisciplines)
                            {
                                double v = baseVariant.ContainsKey(d.UniqueName) ? baseVariant[d.UniqueName] : (current.ContainsKey(d.UniqueName) ? current[d.UniqueName] : 0);
                                if (d.Semester == sem1) sumSem1 += v;
                                if (d.Semester == sem2) sumSem2 += v;
                                if (d.Semester == sem1 || d.Semester == sem2)
                                {
                                    sumPair += v;
                                    if (excludedDisciplineNames.Contains(d.Name)) 
                                        excludedSum += v;
                                }
                            }

                            double sumWithoutExcluded = sumPair - excludedSum;
                            bool ok = true;
                            // ... existing constraint checks (sumWithoutExcluded, Math.Abs(sumSem1 - sumSem2)) ...
                            if (Math.Abs(sumWithoutExcluded - targetSum) > 0.001) ok = false;
                            else if (Math.Abs(sumSem1 - sumSem2) > 6) ok = false;


                            if (!ok)
                            {
                                rejectedCount++;
                                return;
                            }

                            var variant = new Dictionary<string, double>(baseVariant);
                            foreach (var kv in current) variant[kv.Key] = kv.Value;

                            double objSem1 = 0, objSem2 = 0; // These are sums of (workload * significance)
                            foreach (var d in blockDisciplines)
                            {
                                double v = variant.TryGetValue(d.UniqueName, out double val) ? val : 0;
                                if (d.Semester == sem1) objSem1 += v * d.SignificanceCoefficient;
                                if (d.Semester == sem2) objSem2 += v * d.SignificanceCoefficient;
                            }

                            lock (localVariants)
                            {
                                localVariants.Add(new BlockVariant
                                {
                                    Values = variant,
                                    SumsBySemester = new[] { objSem1, objSem2 }
                                });
                            }
                            return;
                        }
                        var disc = variableBlock[idx];
                        foreach (var value in valueOptions[idx])
                        {
                            current[disc.UniqueName] = value;
                            Recurse(idx + 1, current);
                            current.Remove(disc.UniqueName);
                        }
                    }

                    Recurse(0, new Dictionary<string, double>());
                    Console.WriteLine($"Блок {sem1}+{sem2}: найдено валидных вариантов до фильтрации {localVariants.Count}");

                    // Heuristic pruning for the block based on significance sums (internal objective)
                    var prunedBlockVariants = localVariants
                        .OrderByDescending(v => v.SumsBySemester[0] + v.SumsBySemester[1]) // Simple sum, not product, or another heuristic
                        .Take(maxVariantsPerBlock)
                        .ToList();

                    Console.WriteLine($"Блок {sem1}+{sem2}: отобрано топ-{maxVariantsPerBlock} вариантов");
                    lock (blockVariantsLock)
                    {
                        blockVariants.Add(prunedBlockVariants);
                    }
                });

                var allCombinedResults = new ConcurrentBag<Dictionary<string, double>>();
                int blockCount = blockVariants.Count;

                if (blockCount == 0 || blockVariants.Any(b => b.Count == 0))
                {
                    Console.WriteLine("Один из блоков не сгенерировал вариантов. Итоговый результат будет пустым.");
                    return new List<Dictionary<string, double>>();
                }

                var blockLengths = blockVariants.Select(b => b.Count).ToArray();
                long totalComb = blockLengths.Aggregate(1L, (a, b) => a * b);
                Console.WriteLine($"Итоговое число комбинаций для объединения: {totalComb}");
                if (totalComb == 0)
                {
                     return new List<Dictionary<string, double>>();
                }

                int maxFinalVariantsToReturn = 100; // Max variants to return to UI

                // Iterate through combinations and collect valid ones
                // No global objective calculation or sorting by it.
                // The Partitioner helps parallelize iterating 'totalComb'
                Parallel.ForEach(Partitioner.Create(0L, totalComb), range =>
                {
                    for (long idx = range.Item1; idx < range.Item2; idx++)
                    {
                        if (allCombinedResults.Count >= maxFinalVariantsToReturn && totalComb > maxFinalVariantsToReturn) // Optimization: stop if we have enough
                        {
                            // This break is for the inner loop of this specific partition.
                            // Other partitions might still add. For a hard stop, a CancellationToken would be better.
                            // For simplicity, we'll let all partitions run but only add if under limit,
                            // or trim at the end. Let's trim at the end for simplicity here.
                        }

                        var indices = new int[blockCount];
                        long t = idx;
                        for (int i = blockCount - 1; i >= 0; i--)
                        {
                            if (blockLengths[i] == 0) continue; 
                            indices[i] = (int)(t % blockLengths[i]);
                            t /= blockLengths[i];
                        }

                        var combinedVariant = new Dictionary<string, double>();
                        for (int b = 0; b < blockCount; b++)
                        {
                            if (indices[b] >= blockVariants[b].Count) continue; 
                            var blockVar = blockVariants[b][indices[b]];
                            foreach (var kv in blockVar.Values)
                                combinedVariant[kv.Key] = kv.Value;
                        }
                        allCombinedResults.Add(combinedVariant);
                    }
                });
                
                return allCombinedResults.Take(maxFinalVariantsToReturn).ToList();
            });
        }

        private static List<double> GeneratePossibleValues(double min, double max)
        {
            var values = new List<double>();
            for (double v = min; v <= max; v += 1) // Assuming step is 1
            {
                values.Add(v);
            }
            return values;
        }

        public static List<Discipline>[] PrepareSemDisc(List<Discipline> disciplines)
        {
            var semDisc = new List<Discipline>[8];
            for (int sem = 1; sem <= 8; sem++)
                semDisc[sem - 1] = disciplines.Where(d => d.Semester == sem).ToList();
            return semDisc;
        }
    }
}
