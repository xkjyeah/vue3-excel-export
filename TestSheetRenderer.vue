<template>
  <SheetJsOutput>
    <template v-slot:sheet>
      <row>
        <text>Employee ID</text>
        <text>Name</text>
        <text>Date Joined</text>
        <text>YoE at Hire</text>
        <text>YoE to date</text>
        <text>Is Manager?</text>
        <text>Created</text>
        <text>Updated</text>
      </row>
      <row v-for="e in employees" :key="e.id" :widthSetting="true">
        <text>{{ e.id }}</text>
        <text>{{ e.name }}</text>
        <date width="11" z="DD MMM YYYY">{{ e.dateJoined }}</date>
        <number z="#,##0.0">{{ e.yoeHire }}</number>
        <formula>
          <rc c="-1" /> + (now() - <rc c="-2" />) / 365
        </formula>
        <boolean>{{ e.isManager }}</boolean>
        <date width="20" z="YYYY-MM-DD HH:MM:SS">{{ e.created }}</date>
        <date width="20" z="YYYY-MM-DD HH:MM:SS">{{ e.created }}</date>
      </row>
    </template>
    <template v-slot:default="output">
      {{ JSON.stringify(output) }}
    </template>
  </SheetJsOutput>
</template>
<script setup>
import SheetJsOutput from './SheetJsOutput'

const employees = [
  {
    id: 101,
    name: 'Matthew',
    dateJoined: '2022-06-01',
    yoeHire: 5.5,
    isManager: false,
    created: '2022-05-30T16:09:10Z',
    updated: '2024-01-01:T01:02:03Z',
  },
  {
    id: 102,
    name: 'Julius',
    dateJoined: '2024-03-05',
    yoeHire: 3.0,
    isManager: true,
    created: '2024-03-01T07:09:10Z',
    updated: '2025-01-01:T01:05:03Z',
  },
]
</script>