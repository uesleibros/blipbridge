$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
New-Item -ItemType Directory -Force "$root/docs/evidence","$root/benchmarks/results" | Out-Null
foreach($name in 'functional_inproc.txt','memory_functional.txt','com_smoke.txt','userpicture_trace.txt','memory_trace.txt','pdb_probe_mso20.txt','pdb_probe_oart.txt','decoder_watch.txt','decoder_stream.txt','cached_factory.txt','cached_apply_trace.txt','gfx_stream_exports.txt','stress_reopen.txt','fill_transaction_validated.txt','fill_transaction_lifecycle.txt','fill_transaction_lifecycle_repeat.txt','decoder_shape_validation.txt','oart_consumer.txt','oart_lifetime_phases.txt','oart_retention.txt','decoder_lifetime_reopen.txt','fallback_contract.txt','receiver_identity.txt','receiver_identity_validation.txt','receiver_lookup.txt','cached_image_load.txt','native_apply.txt','native_apply_reuse.txt','native_apply_stability.txt','native_texture.txt','native_texture_shutdown.txt','texture_benchmark.txt','undo_harness.txt','refcount_attribution.txt','native_texture_stress.txt','no_temp_image.txt','batch_benchmark.txt','abi_in_powerpoint.txt','pixel_formats.txt','pixel_benchmark.txt','slowdown_attribution.txt','stage_profile.txt','shape_matrix.txt','shape_class_stress.txt','connector_isolation.txt','shape_discriminator.txt','picture_cache.txt','semantic_guards.txt','shape_lifecycle_cache.txt','group_fill_propagation.txt','shape_identity.txt','cache_ownership.txt','release_stress.txt'){
 $source=Join-Path "$root/artifacts" $name
 if(Test-Path -LiteralPath $source){Copy-Item -LiteralPath $source -Destination "$root/docs/evidence/$name" -Force}
}
foreach($name in 'baseline.csv','baseline_inproc.csv','focused.csv','stress.csv'){
 $source=Join-Path "$root/artifacts" $name
 if(Test-Path -LiteralPath $source){Copy-Item -LiteralPath $source -Destination "$root/benchmarks/results/$name" -Force}
}
foreach($name in 'powerpoint_typelib','mso_typelib'){
 $source="$root/artifacts/$name.txt"
 if(Test-Path -LiteralPath $source){
  $lines=Get-Content $source
  $active=$false
  $selected=foreach($line in $lines){if($line -match '^TYPE '){$active=$line -match '^TYPE (FillFormat|Shape|ShapeRange|Shapes|PictureFormat) '};if($active){$line}}
  $selected | Set-Content "$root/docs/evidence/${name}_shapes.txt"
 }
}
