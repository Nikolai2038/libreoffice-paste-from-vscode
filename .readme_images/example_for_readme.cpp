  // ========================================
  // Add some constraints to the minimum/maximum pitch values
  // ========================================
  // Make sure that when pitch is out of bounds, screen doesn't get flipped
  if (constrainPitch) {
    if (this->pitch > 89.0f) {
      this->pitch = 89.0f;
    } else if (this->pitch < -89.0f) {
      this->pitch = -89.0f;
    }
  }
  // ========================================