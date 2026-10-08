## v1.1 Beta

모듈 구조를 다시 정리하고 지원 기능을 크게 넓혔습니다.

> 모델 함수는 이제 바로 보내지 않고 모아 둡니다. 마지막에 `ModelCreate` 를 한 번 불러야 CIVIL NX에 반영됩니다. `RunAnalysis` 나 저장 함수로 끝나는 매크로는 그대로 동작합니다.

### 새 기능

- 시공단계, 긴장재, 이동하중, 시간의존 재료
- 해석 제어, 응답스펙트럼, 시간이력, 수화열
- 합성·변단면·값 단면, 기타 요소(트러스·인장·압축·솔리드·벽), 그룹
- 위치 기반 모델링 (`NodeAt`, `BeamSE`, `SelectBox` 등)

### 바뀐 점

- 그룹과 하중조건은 이름을 쓰면 자동으로 만들어집니다.
- `Support` 가 `"fix"`, `"pin"`, `"roller"` 를 받고, `BeamLoad` 가 부재 배열을 받습니다.
- 재질 감쇠비 기본값이 0.05로 바뀌었습니다.

### 없어진 함수

| 없어진 함수 | 대신 쓸 것 |
| --- | --- |
| `Db…`, `ApiSend`, `ApiJson`, `MapiCommand…` | `CallGet` / `CallPut` / `CallPost`, `StorePut` |
| `DefineBoundaryCombination`, `AssignBoundaryCombination` | `BoundaryChange` |

### 받을 파일

- `CivilVBA.zip` - `CivilVBA.bas` + `JsonConverter.bas` (둘 다 가져와야 합니다)
- `CivilVBA.bas` - 모듈만
