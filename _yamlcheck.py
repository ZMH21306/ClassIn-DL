import yaml
with open('.github/workflows/release.yml', encoding='utf-8') as f:
    wf = yaml.safe_load(f)
print('Jobs:', list(wf['jobs'].keys()))
for job_name, job in wf['jobs'].items():
    print(f'  {job_name}: runs-on={job.get("runs-on")}, needs={job.get("needs")}')
    if 'strategy' in job:
        print(f'    matrix: {job["strategy"]["matrix"]}')
    if 'condition' in job:
        print(f'    condition: {job["condition"]}')
    steps = job.get('steps', [])
    for i, step in enumerate(steps):
        name = step.get('name', step.get('uses', step.get('run', '?')[:50]))
        print(f'    step {i}: {name}')
