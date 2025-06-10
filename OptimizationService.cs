using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace Optimizations
{
    public class OptimizationService
    {
        public static Task<List<Dictionary<string, double>>> GenerateValidVariantsAsync(List<Discipline> disciplines)
        {
            return Task.Run(() =>
            {
                var semesterPairs = new List<(int Sem1, int Sem2, double TargetSum)>
                {
                    (1, 2, 60.5),
                    (3, 4, 59.5),
                    (5, 6, 60),
                    (7, 8, 60)
                };

                var blockVariants = new List<List<BlockVariant>>(semesterPairs.Count);
                var blockVariantsLock = new object();

                Parallel.ForEach(semesterPairs, semesterPair =>
                {
                    var (firstSemester, secondSemester, targetSum) = semesterPair;
                    var blockDisciplines = disciplines.Where(discipline => discipline.Semester == firstSemester || discipline.Semester == secondSemester).ToList();
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: дисциплин {blockDisciplines.Count}");
                    var fixedDisciplines = blockDisciplines.Where(discipline => Math.Abs(discipline.MinWorkload - discipline.MaxWorkload) < 0.001).ToList();
                    var variableDisciplines = blockDisciplines.Where(discipline => Math.Abs(discipline.MinWorkload - discipline.MaxWorkload) > 0.001).ToList();
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: фиксированных {fixedDisciplines.Count}, переменных {variableDisciplines.Count}");

                    var baseVariant = fixedDisciplines.ToDictionary(discipline => discipline.UniqueName, discipline => discipline.MinWorkload);
                    var valueOptionsList = variableDisciplines
                        .Select(discipline => GeneratePossibleValues(discipline.MinWorkload, discipline.MaxWorkload))
                        .ToList();
                    if (variableDisciplines.Count > 0)
                        Console.WriteLine($"Блок {firstSemester}+{secondSemester}: всего комбинаций {valueOptionsList.Select(optionList => optionList.Count).Aggregate(1, (accumulator, count) => accumulator * count)}");

                    var localVariants = new List<BlockVariant>();
                    int rejectedVariantsCount = 0;

                    void RecurseVariants(int disciplineIndex, Dictionary<string, double> currentVariant)
                    {
                        if (disciplineIndex == variableDisciplines.Count)
                        {
                            var excludedDisciplineNames = new HashSet<string> { "Онтологическое моделирование", "Проектирование пользовательского интерфейса" };
                            double sumFirstSemester = 0, sumSecondSemester = 0, sumPair = 0, excludedSum = 0;
                            foreach (var discipline in blockDisciplines)
                            {
                                double workloadValue = baseVariant.ContainsKey(discipline.UniqueName) ? baseVariant[discipline.UniqueName] : (currentVariant.ContainsKey(discipline.UniqueName) ? currentVariant[discipline.UniqueName] : 0);
                                if (discipline.Semester == firstSemester) sumFirstSemester += workloadValue;
                                if (discipline.Semester == secondSemester) sumSecondSemester += workloadValue;
                                if (discipline.Semester == firstSemester || discipline.Semester == secondSemester)
                                {
                                    sumPair += workloadValue;
                                    if (excludedDisciplineNames.Contains(discipline.Name))
                                        excludedSum += workloadValue;
                                }
                            }

                            double sumWithoutExcluded = sumPair - excludedSum;
                            bool isValidVariant = true;
                            string rejectionReason = "";
                            if (Math.Abs(sumWithoutExcluded - targetSum) > 0.001)
                            {
                                isValidVariant = false;
                                rejectionReason = $"Сумма (без исключённых) {sumWithoutExcluded} != {targetSum}";
                            }
                            else if (Math.Abs(sumFirstSemester - sumSecondSemester) > 6)
                            {
                                isValidVariant = false;
                                rejectionReason = $"Разница между семестрами {Math.Abs(sumFirstSemester - sumSecondSemester)} > 6";
                            }

                            if (!isValidVariant)
                            {
                                if (rejectedVariantsCount < 5)
                                {
                                    var allWorkloadValues = new Dictionary<string, double>(baseVariant);
                                    foreach (var keyValue in currentVariant) allWorkloadValues[keyValue.Key] = keyValue.Value;
                                    string workloadValuesText = string.Join(", ", allWorkloadValues.Select(keyValue => $"{keyValue.Key}:{keyValue.Value}"));
                                    rejectedVariantsCount++;
                                }
                                return;
                            }

                            var validVariant = new Dictionary<string, double>(baseVariant);
                            foreach (var keyValue in currentVariant) validVariant[keyValue.Key] = keyValue.Value;

                            double objectiveFirstSemester = 0, objectiveSecondSemester = 0;
                            foreach (var discipline in blockDisciplines)
                            {
                                double workloadValue = validVariant[discipline.UniqueName];
                                if (discipline.Semester == firstSemester) objectiveFirstSemester += workloadValue * discipline.SignificanceCoefficient;
                                if (discipline.Semester == secondSemester) objectiveSecondSemester += workloadValue * discipline.SignificanceCoefficient;
                            }

                            lock (localVariants)
                            {
                                localVariants.Add(new BlockVariant
                                {
                                    Values = validVariant,
                                    SumsBySemester = new[] { objectiveFirstSemester, objectiveSecondSemester }
                                });
                            }
                            return;
                        }
                        var currentDiscipline = variableDisciplines[disciplineIndex];
                        foreach (var workloadValue in valueOptionsList[disciplineIndex])
                        {
                            currentVariant[currentDiscipline.UniqueName] = workloadValue;
                            RecurseVariants(disciplineIndex + 1, currentVariant);
                            currentVariant.Remove(currentDiscipline.UniqueName);
                        }
                    }

                    RecurseVariants(0, new Dictionary<string, double>());
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: найдено вариантов {localVariants.Count}");

                    var topTenVariants = localVariants
                        .OrderByDescending(variant => variant.SumsBySemester[0] * variant.SumsBySemester[1])
                        .Take(10)
                        .ToList();

                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: отобрано топ-10 вариантов");
                    lock (blockVariantsLock)
                    {
                        blockVariants.Add(topTenVariants);
                    }
                });

                var allResultVariants = new ConcurrentBag<Dictionary<string, double>>();
                int totalBlocks = blockVariants.Count;
                var blockSizes = blockVariants.Select(blockVariantList => blockVariantList.Count).ToArray();
                long totalCombinations = blockSizes.Aggregate(1L, (accumulator, size) => accumulator * size);
                Console.WriteLine($"Итоговое число комбинаций: {totalCombinations}");

                int maxTopVariants = 100;
                var globalTopVariants = new SortedSet<(double objectiveValue, Dictionary<string, double> variant)>(Comparer<(double, Dictionary<string, double>)>.Create((firstItem, secondItem) =>
                {
                    int comparison = firstItem.Item1.CompareTo(secondItem.Item1);
                    if (comparison == 0)
                        return firstItem.Item2.GetHashCode().CompareTo(secondItem.Item2.GetHashCode());
                    return comparison;
                }));
                object topVariantsLock = new object();

                Parallel.ForEach(
                    Partitioner.Create(0L, totalCombinations),
                    () => new SortedSet<(double objectiveValue, Dictionary<string, double> variant)>(Comparer<(double, Dictionary<string, double>)>.Create((firstItem, secondItem) =>
                    {
                        int comparison = firstItem.Item1.CompareTo(secondItem.Item1);
                        if (comparison == 0)
                            return firstItem.Item2.GetHashCode().CompareTo(secondItem.Item2.GetHashCode());
                        return comparison;
                    })),
                    (combinationRange, loopState, threadLocalTopVariants) =>
                    {
                        for (long combinationIndex = combinationRange.Item1; combinationIndex < combinationRange.Item2; combinationIndex++)
                        {
                            var blockIndices = new int[totalBlocks];
                            long tempIndex = combinationIndex;
                            for (int blockIndex = totalBlocks - 1; blockIndex >= 0; blockIndex--)
                            {
                                blockIndices[blockIndex] = (int)(tempIndex % blockSizes[blockIndex]);
                                tempIndex /= blockSizes[blockIndex];
                            }

                            var combinedVariant = new Dictionary<string, double>();
                            double objectiveValue = 1.0;
                            for (int blockIndex = 0; blockIndex < totalBlocks; blockIndex++)
                            {
                                var selectedBlockVariant = blockVariants[blockIndex][blockIndices[blockIndex]];
                                foreach (var workloadPair in selectedBlockVariant.Values)
                                    combinedVariant[workloadPair.Key] = workloadPair.Value;
                                objectiveValue *= (selectedBlockVariant.SumsBySemester[0] + selectedBlockVariant.SumsBySemester[1]);
                            }

                            if (threadLocalTopVariants.Count < maxTopVariants)
                            {
                                threadLocalTopVariants.Add((objectiveValue, combinedVariant));
                            }
                            else if (objectiveValue > threadLocalTopVariants.Min.Item1)
                            {
                                threadLocalTopVariants.Remove(threadLocalTopVariants.Min);
                                threadLocalTopVariants.Add((objectiveValue, combinedVariant));
                            }
                        }
                        return threadLocalTopVariants;
                    },
                    threadLocalTopVariants =>
                    {
                        lock (topVariantsLock)
                        {
                            foreach (var topVariant in threadLocalTopVariants)
                            {
                                if (globalTopVariants.Count < maxTopVariants)
                                {
                                    globalTopVariants.Add(topVariant);
                                }
                                else if (topVariant.Item1 > globalTopVariants.Min.Item1)
                                {
                                    globalTopVariants.Remove(globalTopVariants.Min);
                                    globalTopVariants.Add(topVariant);
                                }
                            }
                        }
                    }
                );

                var topHundredVariants = globalTopVariants.Reverse().Take(100).ToList();
                return topHundredVariants.Select(variantTuple => variantTuple.Item2).ToList();
            });
        }

        private static List<double> GeneratePossibleValues(double minValue, double maxValue)
        {
            var possibleValues = new List<double>();
            for (double currentValue = minValue; currentValue <= maxValue; currentValue += 1)
            {
                possibleValues.Add(currentValue);
            }
            return possibleValues;
        }

        public static List<Discipline>[] PrepareSemDisc(List<Discipline> disciplines)
        {
            var semesterDisciplines = new List<Discipline>[8];
            for (int semesterNumber = 1; semesterNumber <= 8; semesterNumber++)
                semesterDisciplines[semesterNumber - 1] = disciplines.Where(discipline => discipline.Semester == semesterNumber).ToList();
            return semesterDisciplines;
        }

        public static double CalculateObjectiveFast(List<Discipline>[] semesterDisciplines, Dictionary<string, double> workloadVariant)
        {
            return semesterDisciplines.Take(8)
                .Select(semester => semester
                    .Sum(discipline => workloadVariant[discipline.UniqueName] * discipline.SignificanceCoefficient))
                .Aggregate(1.0, (product, semesterSum) => product * semesterSum);
        }

        public class BlockVariant
        {
            public Dictionary<string, double> Values = new Dictionary<string, double>();
            public double[] SumsBySemester = new double[2];
        }
    }
}
