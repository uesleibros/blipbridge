#pragma once
/**
 * @file range_texture.hpp
 * Filling every Shape in one ShapeRange with one private apply.
 *
 * ## Why a ShapeRange and not a list of Shapes
 *
 * The transaction handed to the OART receiver carries no target: the receiver
 * *is* the target. So a transaction holding several Shapes' picture fills cannot
 * exist, and a batch built out of one has to be a different shape of thing - a
 * receiver that stands for several Shapes. `ShapeRange.Fill` is exactly that.
 * It presents the same PPCORE delegating wrapper as `Shape.Fill`, over the same
 * OART FillFormat, over a control block whose receiver covers the whole range,
 * so the existing structural walk accepts it unchanged.
 *
 * This is why the public entry point takes a `ShapeRange` and not an array of
 * Shape pointers. An array would have to be applied to one at a time, which is
 * what `BB_ApplyTextureBatch` already does and what makes it a marshalling
 * convenience rather than a speed-up. The distinction is the whole feature.
 *
 * ## What it costs
 *
 * Measured on 16.0.14334.20848, both legs in process, gate included: 1.876 ms
 * for 32 Shapes against 6.650 ms one at a time, and 5.549 ms for 100 against
 * 20.935 ms. Roughly 0.12 ms fixed plus 0.043 ms per member, against 0.19 ms
 * per member individually. Absolute figures move with machine load; see
 * docs/apply_fast_path.md.
 *
 * ## Undo
 *
 * One range apply is one undo entry, and one Undo reverts the whole fill - the
 * same as Office's own `ShapeRange.Fill.UserPicture`. Counted with a marker
 * Shape on the undo stack at 4, 8 and 16 members and across a range of five
 * different Shape classes. Filling the same Shapes one at a time leaves one
 * entry each, so the per-Shape path is the coarser of the two.
 *
 * ## The gate is the whole risk
 *
 * The fill reaches every member. A Connector in the range would reach the
 * private backend, and a native picture fill on a Connector terminates
 * PowerPoint. So every member is classified before anything internal is touched
 * and one ineligible member refuses the whole range: all or nothing, with
 * nothing applied and the private apply never entered.
 */

// MinGW requires the Windows base types before the Automation declarations.
#include <windows.h>

#include <oleauto.h>

#include <cstdint>

namespace bb::office {

/**
 * Applies the texture @p handle to every member of @p range.
 *
 * @param range A PowerPoint `ShapeRange`, borrowed for this call and never
 *        retained. Not a `Shape`: a Shape cannot answer `Count` and is refused
 *        with a message saying so.
 * @param handle A texture handle from the native store.
 * @return how many member Shapes were filled.
 *
 * Throws bb::Error if the range is empty, if any member has no validated native
 * picture-fill path, or if the apply itself fails. Every member is classified
 * before any internal object is touched, so a refusal leaves the document
 * exactly as it was.
 *
 * A ShapeRange belongs to one slide's `Shapes` collection, so this necessarily
 * operates on one slide. Filling several slides is one call per slide.
 */
std::uint32_t ApplyTextureToRange(IDispatch* range, long handle);

} // namespace bb::office
