# AGENTS.md

Guidance for AI coding agents working in this repository.

## Project overview

A 2D top-down soccer game built with **Godot 4.6** (GL Compatibility renderer), written entirely in **GDScript**. It is a learning/tutorial project (repo name: `Learn-2D-soccer-game-in-Godot`) and is a work in progress: the core gameplay (players, ball, movement, passing, shooting) exists, but menus, CPU AI, and other screens are not implemented yet even though their art assets are already in `assets/`.

There is no package manager, build system, or CI. The "key configuration file" is `project.godot` (engine config). Edit it through the Godot editor when possible rather than by hand.

- Engine: Godot 4.6 (see `config/features` in `project.godot`). The surrounding folder name mentions a mono build, but this project contains **no C# code** — only GDScript (`*.gd`).
- Main scene: `scenes/world.tscn` (`run/main_scene="uid://coar0w52ewb50"`).
- Autoload singleton: `KeyUtils` → `utils/key_utils.gd`.
- Viewport: 280×180 scaled up with `stretch/mode="viewport"` and integer scaling; window override 1400×900.
- Rendering: `gl_compatibility`, nearest-neighbor texture filtering (pixel art).
- Physics layers: 1 `PitchWalls`, 2 `Player`, 3 `Ball`.
- 2-player local input maps are defined in `project.godot` (`p1_*` arrow keys + `[`/`]`, `p2_*` WASD + backtick/`1`). Note the key bindings are unusual — check `project.godot` before assuming.

## Build, run, and test commands

- Open/edit: open the project root in the Godot 4.6 editor.
- Run from CLI: `godot --path .` (or `godot -e` for the editor). The game runs the main scene `scenes/world.tscn`.
- **There are no automated tests, no test framework, and no CI configuration.** Verification is done by running the game in the editor.
- `.godot/` is generated engine cache and is git-ignored; do not edit it.

## Code organization

Two gameplay domains, each with the same architecture: a physics body plus a runtime **state machine** built from plain `Node` children.

```
scenes/
  world.tscn                 # Main scene: pitch background, players, ball (y_sort_enabled)
  ball/
    ball.gd / ball.tscn      # Ball (class_name Ball, AnimatableBody2D) — has fake "height" physics
    ball_state_factory.gd    # Maps Ball.State enum -> state class
    ball_states/
      ball_state.gd          # Base class (class_name BallState): setup(), gravity helper, signals
      ball_state_carried.gd  # Follows the carrier
      ball_state_freeform.gd # Friction, bounce, pickup via PlayerDetectionArea
      ball_state_shot.gd     # Shot behavior
  characters/
    player.gd / player.tscn  # Player (class_name Player, CharacterBody2D)
    player_state_factory.gd  # Maps Player.State enum -> state class
    character_states/
      player_state.gd              # Base class (class_name PlayerState): setup(), transition_state()
      player_state_data.gd         # Builder-style payload passed between states (shot direction/power)
      player_state_moving.gd       # Human input movement; CPU branch is a stub ("process AI movement")
      player_state_passing.gd
      player_state_prepping_shot.gd
      player_state_recovering.gd
      player_state_shooting.gd
      player_state_tackling.gd
utils/
  key_utils.gd               # Autoload: maps Player.ControlScheme -> InputMap actions
assets/
  art/ (backgrounds, characters, palettes, particles, props, ui)   # Pixel-art sprites
  fonts/ (Daydream.ttf, Pixeled.ttf)
  json/squads.json           # Team/player data: country, name, skin, role, speed, power (9 teams)
  music/ (menu, gameplay, tournament, win mp3s)
  sfx/ (bounce, pass, shoot, tackle, whistle, ui wavs)
```

### State machine pattern (read before touching gameplay code)

- Both `Player` and `Ball` hold `current_state` and a `*StateFactory`; states are created with `get_fresh_state()`, wired via `setup(...)` and the `state_transition_requested` signal, then added as a **deferred child node** named `PlayerStateMachine:<state>` / `BallStateMachine`.
- States request transitions by emitting `state_transition_requested`; they never replace themselves directly. Player state transitions may carry a `PlayerStateData` payload (built with `PlayerStateData.build().set_...()`).
- Animation-driven states (e.g. shooting, tackling) override `on_animation_complete()`, which the owner calls when its `AnimationPlayer` finishes; transition happens there, not in `_process`.
- To add a new state: create the class extending `PlayerState`/`BallState`, add the enum value to `Player.State`/`Ball.State`, and register it in the corresponding factory's `states` dictionary.
- Scene nodes are accessed with unique-name references (`%AnimationPlayer`, `%PlayerSprite`, etc.) via `@onready`; when adding nodes to a `.tscn`, mark them as unique names to keep this working.
- `Ball` simulates height manually (`height`, `height_velocity`, gravity in `BallState.process_gravity`) and draws the shadow/offset in `_process`.

### Input

Never read raw keys in gameplay code — go through the `KeyUtils` autoload (`get_input_vector`, `is_action_pressed/just_pressed/just_released`) with the player's `ControlScheme` (`P1`, `P2`, `CPU`). Adding a new action requires an entry in `KeyUtils.Action`, the `ACTIONS_MAP`, and matching `p1_*`/`p2_*` InputMap entries in `project.godot`.

## Code style guidelines

- GDScript, Godot 4.x syntax. All classes use `class_name` and are referenced globally (no explicit `preload` of scripts).
- Typed variables and function signatures are used throughout (`var x : Type`, `func f() -> void`); keep new code typed the same way.
- Indentation: tabs (Godot default). `.editorconfig` only sets `charset = utf-8`; `.gitattributes` enforces `eol=lf` for text files.
- Naming: `snake_case` for files, variables, functions; `PascalCase` for classes; `SCREAMING_SNAKE_CASE` for constants/enums.
- Follow the existing per-domain folder-per-scene layout: each scene lives with its script, its state classes, and its factory.
- Note: some existing identifiers contain typos that are part of the API (`ball.friciton_air`, `context_animation_palyer`). Match the existing spelling when referencing them; renaming is a deliberate refactor, not a drive-by fix.

## Testing instructions

There is no test suite. After gameplay changes, verify by running the project in the Godot editor (F5) and exercising the affected behavior (movement, pass, shoot, tackle, ball pickup/bounce). `project.godot` disables the `unused_signal` GDScript warning — that is intentional (states emit signals connected via `bind()`).

## Security considerations

- No secrets, credentials, or network code in this project; `export_credentials.cfg` is git-ignored (Godot export signing).
- `assets/json/squads.json` is static game data loaded at runtime; it is not validated, so keep edits well-formed (trailing commas are present in the source — verify parsing if you regenerate it).
- Fonts in `assets/fonts/` include a personal-use license file (`Daydream 1.0 Personal License.txt`) — respect it before redistributing or exporting builds.
