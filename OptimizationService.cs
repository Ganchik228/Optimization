using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace Optimizations
{
    public class OptimizationService
    {
        public static Task<List<Dictionary<string, double>>> GenerateValidVariantsAsync(List<Discipline> disciplines, HashSet<string> userExcludedDisciplineNames, List<(int Sem1, int Sem2, double TargetSum)> semesterPairs)
        {
            return Task.Run(() =>
            {
                
                // var semesterPairs = new List<(int Sem1, int Sem2, double TargetSum)>
                // {
                //     (1, 2, 60.5),
                //     (3, 4, 59.5),
                //     (5, 6, 60),
                //     (7, 8, 60)
                // };

                var blockVariants = new List<List<BlockVariant>>(semesterPairs.Count);
                var blockVariantsLock = new object();

                Parallel.ForEach(semesterPairs, semesterPair =>
                {
                    var (firstSemester, secondSemester, targetSum) = semesterPair;
                    var blockDisciplines = disciplines.Where(discipline => discipline.Semester == firstSemester || discipline.Semester == secondSemester).ToList();
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: дисциплин {blockDisciplines.Count}");
                    
                    // Диагностика: проверяем есть ли дисциплины в блоке
                    if (blockDisciplines.Count == 0)
                    {
                        Console.WriteLine($"ВНИМАНИЕ: Блок {firstSemester}+{secondSemester} не содержит дисциплин!");
                        lock (blockVariantsLock)
                        {
                            blockVariants.Add(new List<BlockVariant>());
                        }
                        return;
                    }
                    
                    var fixedDisciplines = blockDisciplines.Where(discipline => Math.Abs(discipline.MinWorkload - discipline.MaxWorkload) < 0.001).ToList();
                    var variableDisciplines = blockDisciplines.Where(discipline => Math.Abs(discipline.MinWorkload - discipline.MaxWorkload) > 0.001).ToList();
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: фиксированных {fixedDisciplines.Count}, переменных {variableDisciplines.Count}");

                    // Диагностика: проверяем сумму фиксированных дисциплин
                    var fixedSum = fixedDisciplines.Sum(d => d.MinWorkload);
                    var minVariableSum = variableDisciplines.Sum(d => d.MinWorkload);
                    var maxVariableSum = variableDisciplines.Sum(d => d.MaxWorkload);
                    var totalMinSum = fixedSum + minVariableSum;
                    var totalMaxSum = fixedSum + maxVariableSum;
                    
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: фиксированная сумма = {fixedSum:F1}");
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: возможный диапазон общей суммы = {totalMinSum:F1} - {totalMaxSum:F1}");
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: целевая сумма = {targetSum:F1}");
                    
                    // Проверяем, возможно ли достичь целевой суммы
                    if (targetSum < totalMinSum - 0.1 || targetSum > totalMaxSum + 0.1)
                    {
                        Console.WriteLine($"КРИТИЧЕСКАЯ ОШИБКА: Целевая сумма {targetSum} недостижима для блока {firstSemester}+{secondSemester}!");
                        Console.WriteLine($"Возможный диапазон: {totalMinSum:F1} - {totalMaxSum:F1}");
                        lock (blockVariantsLock)
                        {
                            blockVariants.Add(new List<BlockVariant>());
                        }
                        return;
                    }

                    var baseVariant = fixedDisciplines.ToDictionary(discipline => discipline.UniqueName, discipline => discipline.MinWorkload);
                    var valueOptionsList = variableDisciplines
                        .Select(discipline => GeneratePossibleValues(discipline.MinWorkload, discipline.MaxWorkload))
                        .ToList();
                    if (variableDisciplines.Count > 0)
                        Console.WriteLine($"Блок {firstSemester}+{secondSemester}: всего комбинаций {valueOptionsList.Select(optionList => optionList.Count).Aggregate(1, (accumulator, count) => accumulator * count)}");

                    var localVariants = new List<BlockVariant>();
                    int rejectedVariantsCount = 0;
                    int totalVariantsChecked = 0;

                    void RecurseVariants(int disciplineIndex, Dictionary<string, double> currentVariant)
                    {
                        if (disciplineIndex == variableDisciplines.Count)
                        {
                            totalVariantsChecked++;
                            var excludedDisciplineNames = userExcludedDisciplineNames; 
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
                            
                            // Увеличиваем допуск для целевой суммы
                            if (Math.Abs(sumWithoutExcluded - targetSum) > 0.5)
                            {
                                isValidVariant = false;
                                rejectionReason = $"Сумма (без исключённых) {sumWithoutExcluded:F1} != {targetSum:F1} (разница: {Math.Abs(sumWithoutExcluded - targetSum):F1})";
                            }
                            else if (Math.Abs(sumFirstSemester - sumSecondSemester) > 6)
                            {
                                isValidVariant = false;
                                rejectionReason = $"Разница между семестрами {Math.Abs(sumFirstSemester - sumSecondSemester):F1} > 6";
                            }

                            if (!isValidVariant)
                            {
                                rejectedVariantsCount++;
                                if (rejectedVariantsCount <= 10)
                                {
                                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: Отклонен вариант - {rejectionReason}");
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
                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: проверено вариантов {totalVariantsChecked}, найдено подходящих {localVariants.Count}, отклонено {rejectedVariantsCount}");

                    var topTenVariants = localVariants
                        .OrderByDescending(variant => variant.SumsBySemester[0] * variant.SumsBySemester[1])
                        .Take(10)
                        .ToList();

                    Console.WriteLine($"Блок {firstSemester}+{secondSemester}: отобрано топ-{topTenVariants.Count} вариантов");
                    lock (blockVariantsLock)
                    {
                        blockVariants.Add(topTenVariants);
                    }
                });

                var allResultVariants = new ConcurrentBag<Dictionary<string, double>>();
                int totalBlocks = blockVariants.Count;
                var blockSizes = blockVariants.Select(blockVariantList => blockVariantList.Count).ToArray();
                
                if (blockSizes.Any(size => size == 0))
                {
                    Console.WriteLine("Один или несколько блоков не имеют допустимых вариантов. Возвращаем пустой список.");
                    return new List<Dictionary<string, double>>();
                }
                
                long totalCombinations = blockSizes.Aggregate(1L, (accumulator, size) => accumulator * size);
                Console.WriteLine($"Итоговое число комбинаций: {totalCombinations}");

                if (totalCombinations <= 0)
                {
                    Console.WriteLine("Общее количество комбинаций равно 0 или отрицательно. Возвращаем пустой список.");
                    return new List<Dictionary<string, double>>();
                }

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
