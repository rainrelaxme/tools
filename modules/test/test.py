from typing import List


class Solution:
    def twoSum(self, nums: List[int], target: int) -> List[int]:
        for num, idx in enumerate(nums):
            for b, index in enumerate(nums):
                if index == idx:
                    continue
                if num + b == target:
                    return [idx, index]


sln = Solution()
print(sln.twoSum([2,7,11,15], 9))